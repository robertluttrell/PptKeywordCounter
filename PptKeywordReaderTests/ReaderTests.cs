using System;
using Xunit;
using PptReader;

namespace PptReaderTests
{
    public class ReaderTests
    {
        [Fact]
        public void Test1()
        {
            Reader reader = new Reader(new string[] { "asdf" });
            reader.CountKeywordsAllFiles();
            Console.WriteLine("Debug boi");
        }
    }
}
