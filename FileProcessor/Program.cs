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
            Console.WriteLine("Enter path to directory containing Powerpoint files:");
            string pptPath = Console.ReadLine();

            Console.WriteLine("Enter Excel destination path:");
            string outputPath = Console.ReadLine();

            List<string> filePaths = Directory.GetFiles(pptPath).ToList().Where(s => s.EndsWith(".pptx")).ToList();

            Processor p = new Processor(filePaths, outputPath);
            p.ProcessFiles();
        }
    }
}
