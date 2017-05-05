using System;

namespace XmlParserSample
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Parsing intiated..\n");
            LoadAndManipulateXml.LoadXmlFilesAndUpdate();
            Console.WriteLine("Parsing completed successfully.\n\nPress any key to continue...");
            Console.ReadKey();
        }
    }
}
