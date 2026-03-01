// See https://aka.ms/new-console-template for more information
using ExcelConsolidator;
using ExcelConsolidator.Models;
using ExcelConsolidator.Services;
using INIParser;

class Program_old
{
    static void Main_old(string[] args)
    {
        var hasArgs = CheckArgs(args);

        if (hasArgs)
        {
            var iniFile = new IniFile();
            //iniFile.LoadFile(@"C:\Users\jvand\source\repos\ExcelConsolidator\config.ini");
            iniFile.LoadFile(args[0]);

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

    private static bool CheckArgs(string[] args)
    {
        // Check if any arguments were passed
        if (args.Length > 0)
        {
            //Console.WriteLine($"Total arguments: {args.Length}");

            //foreach (string arg in args)
            //{
            //    Console.WriteLine($"Argument: {arg}");
            //}
            return true;
        }
        else
        {
            Console.WriteLine("You didn't add a path to the configuration file" +
                "\nCorrect execution of this app looks like: " +
                "\nExcelConsolidator \"c:\\mydirectory\\config.ini\"");

            return false;
        }
    }
}
