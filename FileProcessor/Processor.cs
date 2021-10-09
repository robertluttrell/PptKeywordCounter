using System;
using System.Collections.Generic;
using FileOccurrence;
using PowerpointReader;
using ExcelWriter;

namespace FileProcessor
{
    public class Processor
    {

        private readonly List<string> _filePaths;  // List of powerpoint filepaths to be read
        private readonly string _outputPath;  // Output filepath for excel file
        private List<KeywordFileOccurrence> _kfoList;  // List of file occurrences created by the PptReader to be passed to the ExcelWriter

        public Processor(List<string> filePaths, string outputPath)
        {
            _filePaths = filePaths;
            _outputPath = outputPath;
        }

        /// <summary>
        /// Reads keywords from all filepaths and writes output to an excel spreadsheet at _outputPath
        /// </summary>
        public void ProcessFiles()
        {
            // Read from powerpoint files
            PptReader reader = new PptReader(_filePaths);
            reader.CountKeywordsAllFiles();

            // Access and sort file occurrence list from reader to pass to writer
            _kfoList = reader.kfoList;
            _kfoList.Sort(new KeywordFileOccurrenceComparer());

            // Write to excel file
            ExcelWriter.ExcelWriter writer = new ExcelWriter.ExcelWriter(_outputPath, _kfoList);
            writer.CreateAndWriteSpreadsheet();
        }
    }
}
