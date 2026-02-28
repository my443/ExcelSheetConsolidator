// See https://aka.ms/new-console-template for more information
class Program
{
    static void Main(string[] args)
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
