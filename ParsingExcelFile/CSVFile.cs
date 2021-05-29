using CsvHelper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ParsingExcelFile
{
    public class CSVFile
    {
        public string Path { get; set; }
        public ExcelFile Excel { get; set; }
        public CSVFile()
        {
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                Path = dialog.SelectedPath + "\\parsed_excel-file.csv";

                if (File.Exists(Path)) File.Delete(Path);

                Excel = new ExcelFile(Path);
            }
        }
        public void WriteField(int row, int col, string input)
        {
            Excel[row,col] = input;
        }
    }
}
