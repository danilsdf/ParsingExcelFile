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
        public StreamWriter Writer { get; set; }
        public CsvWriter CsvWriter { get; set; }
        public CSVFile()
        {
            var dialog = new Ookii.Dialogs.Wpf.VistaFolderBrowserDialog();
            if (dialog.ShowDialog().GetValueOrDefault())
            {
                Path = dialog.SelectedPath + "\\parsed_excel-file.csv";

                if (File.Exists(Path)) File.Delete(Path);

                Writer = new StreamWriter(new FileStream(Path, FileMode.CreateNew), Encoding.UTF8);
                CsvWriter = new CsvWriter(Writer, CultureInfo.InvariantCulture);
            }
        }
        public void WriteField(string input)
        {
            CsvWriter.WriteField(input);
            NextRecord();
            Writer.Flush();
        }
        public void NextRecord()
        {
            CsvWriter.NextRecord();
            Writer.Flush();
        }
    }
}
