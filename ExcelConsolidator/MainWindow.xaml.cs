using ExcelConsolidator.Models;
using ExcelConsolidator.Services;
using INIParser;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace ExcelConsolidator
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
            var iniFile = new IniFile();
            iniFile.LoadFile(@"C:\Users\jvand\source\repos\ExcelConsolidator\config.ini");

            // Retrieve the values using the Section and Key names
            string folderPath = iniFile["Paths", "SourceFolder"];
            string templateFilePath = iniFile["Paths", "TemplateFile"];
            string outputFilePath = iniFile["Paths", "OutputFile"];

            //string folderPath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\Directory Of Files";
            //string templateFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\SampleTemplate.xlsx";
            //string outputFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\OutputFile.xlsx";

            var extractionTempalte = new ExtractionTemplate();

            ExportTemplate template = extractionTempalte.GetTemplateItems(templateFilePath);
            ExcelExtraction extraction = new ExcelExtraction(template);
            ExportRowsCollection rowsCollection = extraction.ExtractDataFromDirectory(folderPath);

            ExcelExport excelExport = new ExcelExport(outputFilePath, rowsCollection, template);
        }
    }
}
