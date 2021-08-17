using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;

namespace PptReader
{
    public class Reader
    {
        private readonly string[] _filePaths;

        public Reader(string[] filePaths)
        {
            _filePaths = filePaths;
            _filePaths = new string[] { @"C:\Users\rober\source\repos\PptKeywordReader\TestFiles\Presentation1.pptx" };
            KeywordDict = new Dictionary<string, List<KeywordFileOccurrence>>();
        }

        public void AddPresentationKeywordsToMasterDict(Dictionary<string, List<int>> presentationKeywordDict, string filePath)
        {
            foreach (string keyword in presentationKeywordDict.Keys)
            {
                if (!KeywordDict.ContainsKey(keyword))
                {
                    KeywordDict.Add(keyword, new List<KeywordFileOccurrence> { new KeywordFileOccurrence(keyword, filePath, presentationKeywordDict[keyword]) });
                }

                else
                {
                    KeywordDict[keyword].Add(new KeywordFileOccurrence(keyword, filePath, presentationKeywordDict[keyword]));
                }
            }
        }

        public void CountKeywordsAllFiles()
        {
            foreach (string filePath in _filePaths)
            {
                var nestedKeywordList = CountKeywordsSingleFile(filePath);
                var presentationKeywordDict = MakePresentationKeywordDict(nestedKeywordList, filePath);
                AddPresentationKeywordsToMasterDict(presentationKeywordDict, filePath);
            }
        }

        private Dictionary<string, List<int>> MakePresentationKeywordDict(List<string> presentationKeywordList, string filePath)
        {
            var presentationKeywordDict = new Dictionary<string, List<int>>();

            for (int slideIndex = 0; slideIndex < presentationKeywordList.Count(); slideIndex++)
            {
                string keywordListRaw = presentationKeywordList[slideIndex];
                if (keywordListRaw != null)
                {
                    List<string> slideKeywordList = keywordListRaw.Replace("Keywords:", "").Trim().Split(",").ToList();
                    foreach (string keyword in slideKeywordList.ConvertAll(d => d.ToLower()))
                    {
                        if (!presentationKeywordDict.ContainsKey(keyword))
                        {
                            presentationKeywordDict.Add(keyword, new List<int> { slideIndex });
                        }
                        else
                        {
                            presentationKeywordDict[keyword].Add(slideIndex);
                        }
                    }
                }
            }

            return presentationKeywordDict;


        }

        private List<string> CountKeywordsSingleFile(string filePath)
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
                    string s = shape.TextBody.InnerText;
                    if (s.StartsWith("Keywords"))
                    {
                        return s;
                    }
                }
            }

            return null;

        }



        public Dictionary<string, List<KeywordFileOccurrence>> KeywordDict { get; set; }

    }
}