using Microsoft.Win32;
using System;
using System.IO;

namespace ParsingExcelFile
{
    public class UploadExcelFile
    {
        public void YourMethod()
        {
            var openFileDialog = new OpenFileDialog();

            openFileDialog.InitialDirectory = "C:\\Users\\Admin\\Downloads\\Telegram Desktop\\";
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog.RestoreDirectory = true;
            //Open the Pop-Up Window to select the file 
            if (openFileDialog.ShowDialog() == true)
            {
                using Stream s = openFileDialog.OpenFile();
                TextReader reader = new StreamReader(s);
                string st = reader.ReadToEnd();
                var path = openFileDialog.FileName;
                ExcelFile excel = new ExcelFile(path);
                var text = string.Empty;
                CSVFile csvFile = new CSVFile();
                for (int i = 1; i <= excel.GetRows(); i++)
                {
                    for (int j = 1; j <= excel.GetColumns(); j++)
                    {
                        var str = excel[i, j];
                        csvFile.WriteField(i, j, str);
                    }
                }
            }
        }
    }
}
