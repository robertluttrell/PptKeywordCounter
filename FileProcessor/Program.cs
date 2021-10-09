using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace FileProcessor
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.Error.WriteLine("Invalid arguments");
                Console.Error.WriteLine("Usage: KeywordCounter.exe <PowerpointDirectory> <ExcelOutputPath>");
                return;
            }

            string pptPath = args[0];
            string outputPath = args[1];

            List<string> filePaths = Directory.GetFiles(pptPath).ToList().Where(s => s.EndsWith(".pptx")).ToList();

            Processor p = new Processor(filePaths, outputPath);
            p.ProcessFiles();
        }
    }
}
