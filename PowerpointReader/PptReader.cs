using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using FileOccurrence;

namespace PowerpointReader
{
    public class PptReader
    {
        private readonly List<string> _filePaths;

        public PptReader(List<string> filePaths)
        {
            _filePaths = filePaths;
            kfoList = new List<KeywordFileOccurrence>();
        }

        public void CountKeywordsAllFiles()
        {
            foreach (string filePath in _filePaths)
            {
                var nestedKeywordList = GetKeywordTextboxContentFromSlides(filePath);

                var keywordDictForFile = GetKeywordDictFromNestedKeywordList(nestedKeywordList);

                var kfoListForFile = GetKFOListFromKeywordDict(keywordDictForFile, filePath);

                kfoList = kfoList.Concat(kfoListForFile).ToList();

            }
        }

        private List<KeywordFileOccurrence> GetKFOListFromKeywordDict(Dictionary<string, List<int>> keywordDictForFile, string filePath)
        {
            var kfoListForFile = new List<KeywordFileOccurrence>();

            foreach (string keyword in keywordDictForFile.Keys)
            {
                var slideIndices = keywordDictForFile[keyword];
                var kfo = new KeywordFileOccurrence(keyword, filePath, slideIndices);
                kfoListForFile.Add(kfo);
            }
            return kfoListForFile;
        }

        private Dictionary<string, List<int>> GetKeywordDictFromNestedKeywordList(List<string> nestedKeywordList)
        {
            var keywordDictForFile = new Dictionary<string, List<int>>();

            for (int slideIdx = 0; slideIdx < nestedKeywordList.Count; slideIdx++)
            {
                var slideKeywordListRaw = nestedKeywordList[slideIdx];

                if (slideKeywordListRaw != null)
                {
                    List<string> slideKeywordList = slideKeywordListRaw.Replace("keywords:", "").Trim().Split(',').Select(s => s.Trim()).ToList();
                    foreach (string keyword in slideKeywordList.ConvertAll(d => d.ToLower()))
                    {
                        if (!keywordDictForFile.ContainsKey(keyword))
                        {
                            keywordDictForFile.Add(keyword, new List<int>() { slideIdx + 1});
                        }
                        else
                        {
                            keywordDictForFile[keyword].Add(slideIdx + 1);
                        }
                    }
                }
            }
            return keywordDictForFile;
        }

        private List<string> GetKeywordTextboxContentFromSlides(string filePath)
        {
            var keywordList = new List<string>();

            using (PresentationDocument presentationDocument = PresentationDocument.Open(filePath, false))
            {
                var presentation = presentationDocument.PresentationPart.Presentation;
                var numSlides = presentation.SlideIdList.Count();

                for (int slideIndex = 0; slideIndex < numSlides; slideIndex++)
                {
                    // Get the collection of slide IDs from the slide ID list.
                    DocumentFormat.OpenXml.OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;

                    // Get the relationship ID of the slide.
                    string slidePartRelationshipId = (slideIds[slideIndex] as SlideId).RelationshipId;

                    // Get the specified slide part from the relationship ID.
                    SlidePart slidePart = (SlidePart)presentation.PresentationPart.GetPartById(slidePartRelationshipId);
                    keywordList.Add(GetKeywordsFromSlide(slidePart, slideIndex));
                }
            }
            return keywordList;
        }

        public static string GetKeywordsFromSlide(SlidePart slidePart, int slideIndex)
        {
            // Verify that the slide part exists.
            if (slidePart == null)
            {
                throw new ArgumentNullException("slidePart");
            }

            // If the slide exists...
            if (slidePart.Slide != null)
            {
                // Iterate through all the paragraphs in the slide.
                foreach (Shape shape in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
                {
                    string s = shape.TextBody.InnerText.ToLower();
                    if (s.StartsWith("keywords"))
                    {
                        return s;
                    }
                }
            }

            return null;

        }

        public List<KeywordFileOccurrence> kfoList { get; set; }
    }
}