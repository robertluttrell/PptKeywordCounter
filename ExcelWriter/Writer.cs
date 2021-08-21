using System;
using System.IO;
using PptReader;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ExcelWriter
{
    public class Writer
    {
        private readonly string _outputPath;
        private readonly Dictionary<string, List<KeywordFileOccurrence>> _keywordDict;
        private SpreadsheetDocument _excelDoc;

        public Writer(string outputPath, Dictionary<string, List<KeywordFileOccurrence>> keywordDict)
        {
            _outputPath = outputPath;
            _keywordDict = keywordDict;
        }

        public void WriteDictToFile()
        {
            CreateDocument();

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(_outputPath, true))
            {
                SheetData sheetData = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

                Row headerRow = MakeHeaderRow(sheetData);
                sheetData.Append(headerRow);

                KeywordFileOccurrence kfo = new KeywordFileOccurrence("keyword", "path/filename", new List<int>() { 1, 2, 3 });
                int rowIndex = 2;
                Row newRow = MakeRow(kfo, rowIndex, spreadsheet);
                sheetData.Append(newRow);
            }
        }

        public Row MakeRow(KeywordFileOccurrence kfo, int rowIndex, SpreadsheetDocument spreadsheet)
        {
            Row row = new Row() { RowIndex = 2 };
            var newCellList = new List<Cell>();

            newCellList.Add(MakeCell(kfo.Keyword, "A" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(Path.GetFileName(kfo.FilePath), "B" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(string.Join(",", kfo.SlideIndices), "C" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(kfo.FilePath, "D" + rowIndex.ToString(), false));

            foreach (Cell newCell in newCellList)
            {
                row.Append(newCell);
            }

            return row;
        }

        public Cell MakeCell(string cellText, string cellIndex, bool bold)
        {
            //create a new inline string cell
            Cell cell = new Cell() { CellReference = cellIndex };
            cell.DataType = CellValues.InlineString;

            //create a run for the bold text
            Run run1 = new Run();
            run1.Append(new Text(cellText));
            //create runproperties and append a "Bold" to them
            RunProperties run1Properties = new RunProperties();

            if (bold)
            { 
                run1Properties.Append(new Bold());
            }

            //set the first runs RunProperties to the RunProperties containing the bold
            run1.RunProperties = run1Properties;

            InlineString inlineString = new InlineString();
            inlineString.Append(run1);

            cell.Append(inlineString);

            return cell;
        }

        public Row MakeHeaderRow(SheetData sheetData)
        {
            Row row = new Row() { RowIndex = 1 };

            var columnHeaderCellIds = new List<string>() { "A1", "B1", "C1", "D1" };
            var headerTextList = new List<string>() { "Keyword", "File Name", "Slide Indices", "File Path" };

            for (int i = 0; i < columnHeaderCellIds.Count; i++)
            {
                Cell cell = MakeCell(headerTextList[i], columnHeaderCellIds[i], true);
                row.Append(cell);
            }

            return row;
        }

        public void CreateDocument()
        {
            // Create document
            _excelDoc = SpreadsheetDocument.Create(_outputPath, SpreadsheetDocumentType.Workbook);

            // Add a new WorkbookPart to the document
            WorkbookPart workbookpart = _excelDoc.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            // Add a WorksheetPart to the WorkbookPart
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add sheets to workbook
            Sheets sheets = _excelDoc.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet()
            {
                Id = _excelDoc.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = 1,
                Name = "Worksheet1"
            };
            sheets.Append(sheet);

            workbookpart.Workbook.Save();

            _excelDoc.Close();
        }
    }


}
