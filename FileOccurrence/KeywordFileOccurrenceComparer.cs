using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileOccurrence
{
    // This class provides an implementation for comparison to be used by functions such as the List.Sort()
    public class KeywordFileOccurrenceComparer : IComparer<KeywordFileOccurrence>
    {
        // Interface method implementation to compare KeywordFileOccurrence objects when sorting list in ascending order
        public int Compare(KeywordFileOccurrence kfo1, KeywordFileOccurrence kfo2)
        {
            // Sort by keyword first
            var result = kfo1.Keyword.CompareTo(kfo2.Keyword);
            if (result != 0) return result;

            // If keywords identical, then sort by creation date
            result = kfo1.CreationDate.CompareTo(kfo2.CreationDate);
            if (result != 0) return result;

            // If creation dates identical, then sort by filename
            result = kfo1.FileName.CompareTo(kfo2.FileName);
            return result;
        }
    }
}
