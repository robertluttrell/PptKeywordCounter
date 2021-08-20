using System;
using Xunit;
using PptReader;
using ExcelWriter;
using System.Collections.Generic;

namespace ExcelWriterTests
{
    public class WriterTests
    {
        private readonly string _baseDirectory = @"C:\Users\rober\source\repos\PptKeywordReader";

        [Fact]
        public void Writer_EmptyDict_BlankFile()
        {
            var emptyDict = new Dictionary<string, List<KeywordFileOccurrence>>();

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new Writer(outputPath, emptyDict);
            writer.WriteDictToFile();
        }
    }
}
