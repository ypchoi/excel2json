using System;
using System.IO;

namespace ExcelToJson
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length == 0)
            {
                Usage();
                return 1;
            }

            string sourcePath = args[0];
            string targetPath = (1 < args.Length) ? args[1] : Path.ChangeExtension(sourcePath, "json");

            if (!Path.HasExtension(targetPath))
            {
                targetPath = Path.ChangeExtension(targetPath, "json");
            }

            Document document = new Document();

            if (!document.Load(sourcePath))
                return 1;

            if (!document.Save(targetPath))
                return 1;

            return 0;
        }

        static void Usage()
        {
            Console.WriteLine("Usage");
            Console.WriteLine("ExelToJson <SourcePath> [<TargetPath>]");
            Console.WriteLine("- Example1) ExcelToJson Test.xlsx Test.json");
            Console.WriteLine("- Example2) ExcelToJson Test.xlsx");
        }
    }
}
