using System.Collections.Generic;
using System.Windows;

namespace ParsingExcelFile
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Result_Label.HorizontalContentAlignment = HorizontalAlignment.Center;
            Result_Label.Content = "Selecting File...";
            var path = UploadExcelFile.GetExcelPath();
            Result_Label.Content = "Working and Parsing file...";
            var excel = UploadExcelFile.CreateExcel(path);
            var objects = UploadExcelFile.GetResult(excel, out List<string> properties);
            Result_Label.Content = "Selecting folder to save...";
            var pathToSave = UploadExcelFile.WriteNewFile(properties, objects);
            Result_Label.FontSize = 10;
            Result_Label.Content = $"File has been saved\nPath: {pathToSave}";

        }
    }
}
