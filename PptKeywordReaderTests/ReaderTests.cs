using System;
using Xunit;
using PptReader;
using System.Collections.Generic;

namespace PptReaderTests
{
    public class ReaderTests
    {
        private readonly string _baseDirectory = @"C:\Users\rober\source\repos\PptKeywordReader";

        [Fact]
        public void Reader_CountKeywordsAllFiles_SingleFileSingleKeyword_KeywordRecorded()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\SingleKeyword.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.KeywordDict);
            Assert.Single(reader.KeywordDict["mykeyword"]);
            Assert.Equal("mykeyword", reader.KeywordDict["mykeyword"][0].Keyword);
            Assert.Equal(presentationPath, reader.KeywordDict["mykeyword"][0].FilePath);
            Assert.Equal(new List<int>() { 1 }, reader.KeywordDict["mykeyword"][0].SlideIndices);

        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_LowercaseKeywordPrefix_KeywordRecorded()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\LowercaseKeywordPrefix.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.KeywordDict);
            Assert.Single(reader.KeywordDict["mykeyword"]);


        }
    }
}
