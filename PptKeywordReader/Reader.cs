using System;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Counter
{
    public class PptReader
    {
        private readonly string[] _filePaths;

        public PptReader(string[] filePaths)
        {
            _filePaths = filePaths;
            _filePaths = new string[] { @"C:\Users\rober\source\repos\PptKeywordCounter-csharp\TestFiles" };
        }

        public void CountKeywordsAllFiles()
        {
            foreach (string filePath in _filePaths)
            {
                CountKeywordsSingleFile(filePath);
            }
        }

        private void CountKeywordsSingleFile(string filePath)
        {
            return;
        }


        public Dictionary<string, List<KeywordFileOccurrence>> KeywordDict { get; set; }

    }
}