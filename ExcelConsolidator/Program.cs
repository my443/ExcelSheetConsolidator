// See https://aka.ms/new-console-template for more information
using ExcelConsolidator;
using ExcelConsolidator.Models;

class Program
{
    static void Main(string[] args)
    {
        CheckArgs(args);        

        string folderPath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\Directory Of Files";
        string templateFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\SampleTemplate.xlsx";
        string outputFilePath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\OutputFile.xlsx";

        var extractionTempalte = new ExtractionTemplate();

        ExportTemplate template = extractionTempalte.GetTemplateItems(templateFilePath);
        ExcelExtraction extraction = new ExcelExtraction(template);
        extraction.ExtractDataFromDirectory(folderPath);

    }

    private static void CheckArgs(string[] args)
    {
        // Check if any arguments were passed
        if (args.Length > 0)
        {
            Console.WriteLine($"Total arguments: {args.Length}");

            foreach (string arg in args)
            {
                Console.WriteLine($"Argument: {arg}");
            }
        }
        else
        {
            Console.WriteLine("No arguments were provided.");
        }
    }
}
