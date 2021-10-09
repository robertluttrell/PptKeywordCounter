using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileOccurrence
{
    public class KeywordFileOccurrence
    {
        public KeywordFileOccurrence(string keyword, string filePath, List<int> slideIndices, DateTime creationDate)
        {
            Keyword = keyword;
            FilePath = filePath;
            FileName = Path.GetFileName(FilePath);
            SlideIndices = slideIndices;
            CreationDate = creationDate;
        }

        public string Keyword { get; set; }
        public string FilePath { get; set; }
        public string FileName { get; set; }
        public List<int> SlideIndices { get; set; }
        public DateTime CreationDate { get; set; }
    }
}