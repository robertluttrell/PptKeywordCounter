using System;
using System.Collections.Generic;
using FileOccurrence;
using PowerpointReader;
using ExcelWriter;

namespace FileProcessor
{
    public class Processor
    {

        private readonly List<string> _filePaths;
        private readonly string _outputPath;
        private Dictionary<string, List<KeywordFileOccurrence>> _keywordDict;

        public Processor(List<string> filePaths, string outputPath)
        {
            _filePaths = filePaths;
            _outputPath = outputPath;
        }

        public void ProcessFiles()
        {
            PptReader reader = new PptReader(_filePaths);
            reader.CountKeywordsAllFiles();

            _keywordDict = reader.KeywordDict;

            ExcelWriter.ExcelWriter writer = new ExcelWriter.ExcelWriter(_outputPath, _keywordDict);
            writer.WriteDictToFile();
        }
    }
}
