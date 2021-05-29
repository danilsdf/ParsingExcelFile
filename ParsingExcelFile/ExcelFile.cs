using OfficeOpenXml;
using System.IO;

namespace ParsingExcelFile
{
    public class ExcelFile
    {
        public string Path { get; set; }
        public FileInfo CurrentFile { get; set; }
        public ExcelFile(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            CurrentFile = new FileInfo(path);
            if (!File.Exists(path)) File.Create(path);

            using ExcelPackage Package = new ExcelPackage(CurrentFile);
            Package.Workbook.Worksheets.Add("Parsed");
            Package.Save();
        }
        public int GetRows()
        {
            using ExcelPackage Package = new ExcelPackage(CurrentFile);
            var Worksheet = Package.Workbook.Worksheets[0];
            var result = Worksheet.Dimension?.Rows ?? 0;
            Package.Save();
            return result;
        }
        public int GetColumns()
        {
            using ExcelPackage Package = new ExcelPackage(CurrentFile);
            var Worksheet = Package.Workbook.Worksheets[0];
            var result = Worksheet.Dimension?.Columns ?? 0;
            Package.Save();
            return result;
        }
        public string this[int row, int col]
        {
            get
            {
                using ExcelPackage Package = new ExcelPackage(CurrentFile);
                var Worksheet = Package.Workbook.Worksheets[0];
                var str = Worksheet.Cells[row, col].Value;

                if (str != null) return str.ToString();

                Package.Save();
                return string.Empty;
            }
            set
            {
                using ExcelPackage Package = new ExcelPackage(CurrentFile);
                var Worksheet = Package.Workbook.Worksheets[0];
                Worksheet.Cells[row, col].Value = value;
                Package.Save();
            }
        }
    }
}
