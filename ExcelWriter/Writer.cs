using System;
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
                AddColumnHeaders(spreadsheet);


            }

        }

        public Cell FillCellBold(string cellText, string cellIndex, SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookPart = spreadsheet.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

            //create a new inline string cell
            Cell cell = new Cell() { CellReference = cellIndex };
            cell.DataType = CellValues.InlineString;

            //create a run for the bold text
            Run run1 = new Run();
            run1.Append(new Text(cellText));
            //create runproperties and append a "Bold" to them
            RunProperties run1Properties = new RunProperties();
            run1Properties.Append(new Bold());
            //set the first runs RunProperties to the RunProperties containing the bold
            run1.RunProperties = run1Properties;

            InlineString inlineString = new InlineString();
            inlineString.Append(run1);

            cell.Append(inlineString);

            return cell;
        }

        public void AddColumnHeaders(SpreadsheetDocument spreadsheet)
        {
            WorkbookPart workbookPart = spreadsheet.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
            Row row = new Row() { RowIndex = 1 };

            var columnHeaderCells = new List<Cell>();
            var columnHeaderCellIds = new List<string>() { "A1", "B1", "C1" };
            var headerTextList = new List<string>() { "header1", "header2", "header3" };

            for (int i = 0; i < columnHeaderCellIds.Count; i++)
            {
                Cell cell = FillCellBold(headerTextList[i], columnHeaderCellIds[i], spreadsheet);
                row.Append(cell);
            }

            sheetData.Append(row);
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
