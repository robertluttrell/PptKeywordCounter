using System;
using Xunit;
using PowerpointReader;
using ExcelWriter;
using FileOccurrence;
using System.Collections.Generic;
using System.IO;

namespace ExcelWriterTests
{
    public class ExcelWriterTests
    {
        private readonly string _baseDirectory = @"C:\Users\rober\Source\repos\PptKeywordCounter";

        [Fact]
        public void Writer_NoKfo_WritesColumnHeaders()
        {
            var kfoList = new List<KeywordFileOccurrence>();
            kfoList.Add(new KeywordFileOccurrence("keyword1", "path/file", new List<int>() { 1, 2, 3 }) );

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new ExcelWriter.ExcelWriter(outputPath, kfoList);
            writer.WriteDictToFile();

            Assert.Equal("Keyword", ExcelReader.GetCellValue(outputPath, "Worksheet1", "A1"));
            Assert.Equal("File Name", ExcelReader.GetCellValue(outputPath, "Worksheet1", "B1"));
            Assert.Equal("Slide Indices", ExcelReader.GetCellValue(outputPath, "Worksheet1", "C1"));
            Assert.Equal("File Path", ExcelReader.GetCellValue(outputPath, "Worksheet1", "D1"));
        }

        [Fact]
        public void Writer_SingleKfo_WritesToFile()
        {
            var kfoList = new List<KeywordFileOccurrence>();
            kfoList.Add(new KeywordFileOccurrence("keyword1", "path/file", new List<int>() { 1, 2, 3 }));

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new ExcelWriter.ExcelWriter(outputPath, kfoList);
            writer.WriteDictToFile();

            Assert.Equal("keyword1", ExcelReader.GetCellValue(outputPath, "Worksheet1", "A2"));
            Assert.Equal("file", ExcelReader.GetCellValue(outputPath, "Worksheet1", "B2"));
            Assert.Equal("1,2,3", ExcelReader.GetCellValue(outputPath, "Worksheet1", "C2"));
            Assert.Equal("path/file", ExcelReader.GetCellValue(outputPath, "Worksheet1", "D2"));
        }

        [Fact]
        public void Writer_MultipleKeywordsOneFile_AlphabeticalByKeyword()
        {
            var kfoList = new List<KeywordFileOccurrence>();
            kfoList.Add(new KeywordFileOccurrence("keyword2", "path/file", new List<int>() { 1, 2, 3 }));
            kfoList.Add(new KeywordFileOccurrence("keyword1", "path/file", new List<int>() { 1, 2, 3 }));

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new ExcelWriter.ExcelWriter(outputPath, kfoList);
            writer.WriteDictToFile();

            Assert.Equal("keyword1", ExcelReader.GetCellValue(outputPath, "Worksheet1", "A2"));
            Assert.Equal("keyword2", ExcelReader.GetCellValue(outputPath, "Worksheet1", "A3"));
        }

        [Fact]
        public void Writer_SingleKeywordMultipleFiles_AlphabeticalByFileName()
        {
            var kfoList = new List<KeywordFileOccurrence>();
            kfoList.Add(new KeywordFileOccurrence("keyword", "path/file2", new List<int>() { 1, 2, 3 }));
            kfoList.Add(new KeywordFileOccurrence("keyword", "path/file1", new List<int>() { 1, 2, 3 }));

            var outputPath = _baseDirectory + @"\testoutput.xlsx";
            var writer = new ExcelWriter.ExcelWriter(outputPath, kfoList);
            writer.WriteDictToFile();

            Assert.Equal("file1", ExcelReader.GetCellValue(outputPath, "Worksheet1", "B2"));
            Assert.Equal("file2", ExcelReader.GetCellValue(outputPath, "Worksheet1", "B3"));
        }
    }
}
