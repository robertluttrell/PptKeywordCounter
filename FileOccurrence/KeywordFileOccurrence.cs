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

        public string Keyword { get; set; }  // Lowercase-converted keyword 
        public string FilePath { get; set; }  // Path to the powerpoint file where this keyword was found
        public string FileName { get; set; }  // The name of the powerpoint file where this keyword was found
        public List<int> SlideIndices { get; set; }  // List of integers representing the slide indices (0-based) where this keyword was found
        public DateTime CreationDate { get; set; }  // Date and time when the powerpoint file with this keyword was created
    }
}