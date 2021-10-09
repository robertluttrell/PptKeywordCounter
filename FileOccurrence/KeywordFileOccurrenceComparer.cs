using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FileOccurrence
{
    public class KeywordFileOccurrenceComparer : IComparer<KeywordFileOccurrence>
    {
        public int Compare(KeywordFileOccurrence kfo1, KeywordFileOccurrence kfo2)
        {
            var result = kfo1.Keyword.CompareTo(kfo2.Keyword);
            if (result != 0) return result;

            result = kfo1.FileName.CompareTo(kfo2.FileName);
            if (result != 0) return result;

            result = kfo1.CreationDate.CompareTo(kfo2.CreationDate);
            return result;
        }
    }
}
