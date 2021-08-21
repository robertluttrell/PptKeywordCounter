using System;
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
            return kfo1.Keyword.CompareTo(kfo2.Keyword);
        }
    }
}
