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
            string basePath = @"C:\Users\rober\source\repos\PptKeywordReader";
            string pptPath = basePath + @"\TestFiles";

            List<string> filePaths = Directory.GetFiles(pptPath).ToList().Where(s => s.EndsWith(".pptx")).ToList();

            string outputPath = basePath + @"\testoutput.xlsx";

            Processor p = new Processor(filePaths, outputPath);
            p.ProcessFiles();
        }
    }
}
