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
        public void Reader_CountKeywordsAllFiles_BlankPresentation_EmptyDict()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\BlankPresentation.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Empty(reader.KeywordDict);
        }

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

        [Fact]
        public void Reader_CountKeywordsAllFiles_TextboxKeyword_CountsKeyword()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\TextboxKeyword.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.KeywordDict);
            Assert.Single(reader.KeywordDict["mykeyword"]);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_BulletPointKeyword_CountsKeyword()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\BulletPointKeyword.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.KeywordDict);
            Assert.Single(reader.KeywordDict["mykeyword"]);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_NewlineBetweenKeywords_CountsKeywords()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\NewlineBetweenKeywords.pptx"; 

            Reader reader = new Reader(new string[] { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Equal(2, reader.KeywordDict.Count);
            Assert.True(reader.KeywordDict.ContainsKey("keyword1"));
            Assert.True(reader.KeywordDict.ContainsKey("keyword2"));
        }
        
        [Fact]
        public void Reader_CountKeywordsAllFiles_MultiplePresentationsMultipleKeywords_CountsKeywords()
        {
            var presentation1Path = _baseDirectory + @"\TestFiles\MultiplePresentationsMultipleKeywords1.pptx"; 
            var presentation2Path = _baseDirectory + @"\TestFiles\MultiplePresentationsMultipleKeywords2.pptx"; 

            Reader reader = new Reader(new string[] { presentation1Path, presentation2Path });
            reader.CountKeywordsAllFiles();

            Assert.Equal(2, reader.KeywordDict.Count);

            Assert.Equal(new List<int>() { 1, 2 }, reader.KeywordDict["keyword1"][0].SlideIndices);
            Assert.Equal(new List<int>() { 2 }, reader.KeywordDict["keyword2"][0].SlideIndices);
        }

    }
}
