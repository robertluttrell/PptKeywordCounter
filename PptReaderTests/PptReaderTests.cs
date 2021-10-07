using System;
using Xunit;
using PowerpointReader;
using FileOccurrence;
using System.Collections.Generic;
using System.Linq;

namespace PptReaderTests
{
    public class PptReaderTests
    {
        private readonly string _baseDirectory = @"C:\Users\rober\Source\repos\PptKeywordCounter";

        [Fact]
        public void Reader_CountKeywordsAllFiles_BlankPresentation_EmptyDict()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\BlankPresentation.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Empty(reader.kfoList);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_SingleFileSingleKeyword_KeywordRecorded()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\SingleKeyword.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.kfoList);
            Assert.Equal("mykeyword", reader.kfoList[0].Keyword);
            Assert.Equal(presentationPath, reader.kfoList[0].FilePath);
            Assert.Equal(new List<int>() { 1 }, reader.kfoList[0].SlideIndices);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_LowercaseKeywordPrefix_KeywordRecorded()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\LowercaseKeywordPrefix.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.kfoList);
            Assert.Equal("mykeyword", reader.kfoList[0].Keyword);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_TextboxKeyword_CountsKeyword()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\TextboxKeyword.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.kfoList);
            Assert.Equal("mykeyword", reader.kfoList[0].Keyword);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_BulletPointKeyword_CountsKeyword()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\BulletPointKeyword.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Single(reader.kfoList);
            Assert.Equal("mykeyword", reader.kfoList[0].Keyword);
        }

        [Fact]
        public void Reader_CountKeywordsAllFiles_NewlineBetweenKeywords_CountsKeywords()
        {
            var presentationPath = _baseDirectory + @"\TestFiles\NewlineBetweenKeywords.pptx";

            PptReader reader = new PptReader(new List<string> { presentationPath });
            reader.CountKeywordsAllFiles();

            Assert.Equal(2, reader.kfoList.Count);
            var keywords = new HashSet<string>() { reader.kfoList[0].Keyword, reader.kfoList[1].Keyword };
            Assert.Contains("keyword1", keywords);
            Assert.Contains("keyword2", keywords);
        }
        
        [Fact]
        public void Reader_CountKeywordsAllFiles_MultiplePresentationsMultipleKeywords_CountsKeywords()
        {
            var presentation1Path = _baseDirectory + @"\TestFiles\MultiplePresentationsMultipleKeywords1.pptx"; 
            var presentation2Path = _baseDirectory + @"\TestFiles\MultiplePresentationsMultipleKeywords2.pptx";

            PptReader reader = new PptReader(new List<string> { presentation1Path, presentation2Path });
            reader.CountKeywordsAllFiles();

            Assert.Equal(4, reader.kfoList.Count);

            var keyword1KFOSublist = reader.kfoList.Where(k => k.Keyword == "keyword1");
            var keyword2KFOSublist = reader.kfoList.Where(k => k.Keyword == "keyword2");
            Assert.Equal(2, keyword1KFOSublist.Count());
            Assert.Equal(2, keyword2KFOSublist.Count());
        }

    }
}
