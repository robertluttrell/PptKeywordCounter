using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Packaging;
using FileOccurrence;

namespace PowerpointReader
{
    public class PptReader
    {
        private readonly List<string> _filePaths;  // File paths to read 

        public PptReader(List<string> filePaths)
        {
            _filePaths = filePaths;
            kfoList = new List<KeywordFileOccurrence>();
        }

        /// <summary>
        /// Counts all keywords in all files in _filePaths list and adds corresponding
        /// KeywordFileOccurrence objects to the kfoList
        /// </summary>
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

        /// <summary>
        /// Creates a list of KeywordFileOccurrence objects from a dictionary of keywords mapped to slide index lists
        /// </summary>
        /// <param name="keywordDictForFile">dictionary of keywords mapped to slide index lists</param>
        /// <param name="filePath">the path to the powerpoint file where the keywords were found</param>
        /// <returns></returns>
        private List<KeywordFileOccurrence> GetKFOListFromKeywordDict(Dictionary<string, List<int>> keywordDictForFile, string filePath)
        {
            var kfoListForFile = new List<KeywordFileOccurrence>();
            var creationDate = File.GetCreationTime(filePath);

            foreach (string keyword in keywordDictForFile.Keys)
            {
                var slideIndices = keywordDictForFile[keyword];
                var kfo = new KeywordFileOccurrence(keyword, filePath, slideIndices, creationDate);
                kfoListForFile.Add(kfo);
            }
            return kfoListForFile;
        }

        /// <summary>
        /// Gets a dictionary representing keyword occurrences in a file from a list of keyword textbox entries
        /// </summary>
        /// <param name="nestedKeywordList">keyword entries for each slide (null if none exist)</param>
        /// <returns>dictionary with keyword as key and list of slide indices as value</string></returns>
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

        /// <summary>
        /// Gets a list of all textbox contents in a powerpoint file where the content starts with "keywords:"
        /// </summary>
        /// <param name="filePath">path to the powerpoit file to be read</param>
        /// <returns>list of all textbox elements as a string starting with "keywords:"</returns>
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

        /// <summary>
        /// Gets the textbox content for the first text element in a single slide that starts with "keywords:"
        /// </summary>
        /// <param name="slidePart">the slidePart representing the slide at slideIndex</param>
        /// <param name="slideIndex">the index of the slide to be processed</param>
        /// <returns></returns>
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

        public List<KeywordFileOccurrence> kfoList { get; set; }  // list of KeywordFileOccurrence objects discovered by PptReader
    }
}