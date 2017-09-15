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

            string sourceIn = args[0];

            string[] sourceFiles = Directory.GetFiles(
                Path.GetDirectoryName(sourceIn),
                Path.GetFileName(sourceIn),
                SearchOption.AllDirectories);

            foreach (string sourcePath in sourceFiles)
            {
                Document document = new Document();
                if (!document.Load(sourcePath))
                    continue;

                string targetPath = Path.ChangeExtension(sourcePath, "json");
                if (!document.Save(targetPath))
                    continue;
            }

            return 0;
        }

        static void Usage()
        {
            Console.WriteLine("Usage");
            Console.WriteLine("ExelToJson <SourcePath>");
            Console.WriteLine("- Example) ExcelToJson.exe Test.xlsx");
            Console.WriteLine("- Example) ExcelToJson.exe C:\\Dir\\*.*");
        }
    }
}
