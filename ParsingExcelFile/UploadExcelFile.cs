using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ParsingExcelFile
{
    public static class UploadExcelFile
    {
        public static string GetExcelPath()
        {
            var openFileDialog = new OpenFileDialog
            {
                InitialDirectory = "C:\\",
                Filter = "Excel files (*.xlsx)|*.xlsx",
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() == true)
            {
                using Stream s = openFileDialog.OpenFile();
                TextReader reader = new StreamReader(s);
                string st = reader.ReadToEnd();
                return openFileDialog.FileName;
            }
            return "";
        }
        public static ExcelFile CreateExcel(string path)
        {
            return new ExcelFile(path);
        }
        public static Dictionary<string, object>[] GetResult(ExcelFile excel, out List<string> properties)
        {
            properties = new List<string>();
            var objects = new Dictionary<string, object>[excel.Rows - 1];
            for (int i = 0; i < objects.Length; i++)
            {
                objects[i] = new Dictionary<string, object>();
            }
            for (int i = 1; i <= excel.Columns;)
            {
                var prop = FirstCharToUpper(excel[1, i]);
                if (prop == "Название_Характеристики")
                {
                    if (excel[1, i + 1] == "Измерение_Характеристики")
                    {
                        for (int j = 2; j <= excel.Rows; j++)
                        {
                            prop = FirstCharToUpper(excel[j, i]);
                            if (string.IsNullOrEmpty(prop)) continue;
                            var index = properties.IndexOf(prop);
                            if (index == -1) properties.Add(prop);
                            var value = excel[j, i + 2];
                            var meassure = excel[j, i + 1];
                            objects[j - 2].Add(prop, value + meassure);
                        }
                        i += 3;
                    }
                    else
                    {
                        for (int j = 2; j <= excel.Rows; j++)
                        {
                            prop = FirstCharToUpper(excel[j, i]);
                            var index = properties.IndexOf(prop);
                            if (index == -1) properties.Add(prop);
                            objects[j - 2].Add(prop, excel[j, i + 1]);
                        }
                        i += 2;
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(prop)) properties.Add(prop);
                    if (prop == "Поисковые_запросы") properties.Add("Brand");
                    for (int j = 2; j <= excel.Rows; j++)
                    {
                        if (string.IsNullOrEmpty(prop)) continue;
                        if (prop == "Currency ID")
                        {
                            var item = excel[j, i] == "UAH" ? 4 : 1;
                            objects[j - 2].Add(prop, item);
                        }
                        else
                        {
                            if (prop == "Поисковые_запросы")
                            {
                                var _prop = "Brand";
                                objects[j - 2].Add(_prop, "SULLIVAN");
                            }
                            objects[j - 2].Add(prop, excel[j, i]);
                        }
                    }
                    i++;
                }
            }
            return objects;
        }

        public static string WriteNewFile(List<string> properties, Dictionary<string, object>[] objects)
        {
            ExcelFile csvFile = new ExcelFile();
            return csvFile.WriteExcelFile(properties, objects);
        }
        private static string FirstCharToUpper(string input)
        {
            return input switch
            {
                null => throw new ArgumentNullException(nameof(input)),
                "" => "",
                "Уникальный_идентификатор" => "Артикул",
                "Название_позиции" => "Product",
                "Название_группы" => "Category",
                "Валюта" => "Currency ID",
                _ => input.First().ToString().ToUpper() + input[1..],
            };
        }
    }
}
