using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace ParsingExcelFile
{
    public class ExcelFile
    {
        public string Path { get; set; }
        public ExcelPackage Package { get; set; }
        public ExcelWorksheet Worksheet { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        public ExcelFile()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog
            {
                UseDescriptionForTitle = true,
                Description = "Select folder to save result file"
            };
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                var t = DateTime.Now;
                var path = dialog.SelectedPath + $"\\parsed_{t.Month}_{t.Day}_{t.Hour}_{t.Minute}_{t.Second}.xlsx";

                if (File.Exists(path)) File.Delete(path);

                CreateInstance(path);
            }
        }
        public ExcelFile(string path) => CreateInstance(path);
        public void CreateInstance(string path)
        {
            Path = path;
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(path);

            Package = new ExcelPackage(file);
            try
            {
                Worksheet = Package.Workbook.Worksheets[0];
            }
            catch (Exception)
            {
                Worksheet = Package.Workbook.Worksheets.Add("Parser");
            }

            Rows = Worksheet.Dimension?.Rows ?? 0;
            Columns = Worksheet.Dimension?.Columns ?? 0;
        }
        public string this[int row, int col]
        {
            get
            {
                var str = Worksheet.Cells[row, col].Value;

                if (str != null) return str.ToString();

                return string.Empty;
            }
            set
            {
                Worksheet.Cells[row, col].Value = value;
            }
        }
        public string WriteExcelFile(List<string> properties, Dictionary<string, object>[] keyValues)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo file = new FileInfo(Path);

            using var package = new ExcelPackage(file);
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Parser");

            for (int i = 0; i < properties.Count; i++)
            {
                worksheet.Cells[1, i + 1].Value = properties[i].Replace('_', ' ');
            }
            for (int i = 0; i < keyValues.Length; i++)
            {
                foreach (var item in keyValues[i])
                {
                    var col = properties.IndexOf(item.Key) + 1;
                    worksheet.Cells[i + 2, col].Value = item.Value;
                }
            }
            package.Save();
            return Path;
        }
    }
}
