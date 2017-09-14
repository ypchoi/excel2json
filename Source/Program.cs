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

            try
            {
                Sheet data = new Sheet();
                if (!data.Load(sourcePath))
                {
                    return 1;
                }

                if (!data.Save(targetPath))
                {
                    return 1;
                }
            }
            catch (DirectoryNotFoundException /*e*/)
            {
                Console.WriteLine("File not exist : {0}", sourcePath);
                return 1;
            }

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
