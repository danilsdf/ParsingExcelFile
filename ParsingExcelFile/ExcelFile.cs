using Microsoft.Win32;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Controls;

namespace ParsingExcelFile
{
    public class ExcelFile
    {
        public string Path { get; set; }
        public ExcelPackage Package { get; set; }
        public ExcelWorksheet Worksheet { get; set; }
        public int Rows { get; set; }
        public int Columns { get; set; }
        public ExcelFile(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            FileInfo file = new FileInfo(path);

            Package = new ExcelPackage(file);

            Worksheet = Package.Workbook.Worksheets[0];

            Rows = Worksheet.Dimension.Rows;
            Columns = Worksheet.Dimension.Columns;
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
    }
}
