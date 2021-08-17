using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptReader
{
    public class KeywordFileOccurrence
    {
        public KeywordFileOccurrence(string keyword)
        {
            Keyword = keyword;
            SlideIndices = new List<string>();
        }

        public string Keyword { get; set; }

        public List<string> SlideIndices { get; set; }
    }
}