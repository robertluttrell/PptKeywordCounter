using System;
using Xunit;
using PptReader;
using ExcelWriter;
using System.Collections.Generic;

namespace ExcelWriterTests
{
    public class ExcelWriterTests
    {
        private readonly string _baseDirectory = @"C:\Users\rober\source\repos\PptKeywordReader";

        [Fact]
        public void Writer_SingleKfo_WritesToFile()
        {
            var kfoDict = new Dictionary<string, List<KeywordFileOccurrence>>();
            kfoDict.Add("keyword1", new List<KeywordFileOccurrence>() { new KeywordFileOccurrence("keyword1", "path/file", new List<int>() { 1, 2, 3 }) });

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new ExcelWriter.ExcelWriter(outputPath, kfoDict);
            writer.WriteDictToFile();
        }
    }
}
