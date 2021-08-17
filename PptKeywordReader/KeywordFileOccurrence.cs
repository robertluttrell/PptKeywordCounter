using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptReader
{
    public class KeywordFileOccurrence
    {
        public KeywordFileOccurrence(string keyword, string filePath, List<int> slideIndices)
        {
            Keyword = keyword;
            FilePath = filePath;
            SlideIndices = slideIndices;
        }

        public string Keyword { get; set; }
        public string FilePath { get; set; }
        public List<int> SlideIndices { get; set; }
    }
}