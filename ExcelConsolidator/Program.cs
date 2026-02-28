// See https://aka.ms/new-console-template for more information
using ExcelConsolidator;
using ExcelConsolidator.Models;

class Program
{
    static void Main(string[] args)
    {
        CheckArgs(args);
        
        string filepath = @"C:\Users\jvand\source\repos\ExcelConsolidator\SampleFiles\SampleTemplate.xlsx";
        var MappingTemplate = new MappingTemplate();

        List<Cell> list = MappingTemplate.GetTemplateItems(filepath);

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
