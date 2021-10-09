using System;
using System.IO;
using FileOccurrence;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Linq;

namespace ExcelWriter
{
    public class ExcelWriter
    {
        private readonly string _outputPath;  // The path to the excel spreadsheet to be created
        private readonly List<KeywordFileOccurrence> _kfoList;  // List of KeywordFileOccurrence objects created by PptReader
        private uint _lastFilledRow;  // The index of the last row populated in the Excel sheet. Next available row = _lastFilledRow + 1
        private SpreadsheetDocument _excelDoc;  // SpreadSheetDocument representing the excel document in memory

        public ExcelWriter(string outputPath, List<KeywordFileOccurrence> kfoList)
        {
            _outputPath = outputPath;
            _kfoList = kfoList;
        }

        /// <summary>
        /// Creates, populates, and fills the excel sheet at _outputpath with data from _kfoList
        /// </summary>
        public void WriteDictToFile()
        {
            CreateDocument();

            // Using statement disposes of file object when .NET garbage collector recognizes it is no longer in use.
            // This will happen at the end of this using block because the AddRowForKFO call uses the file.
            // The spreadsheet is not saved to _outputPath until the implicit spreadsheet.Close() at the end of this block.
            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(_outputPath, true))
            {
                SheetData sheetData = spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();

                Row headerRow = MakeHeaderRow(sheetData);
                sheetData.Append(headerRow);

                foreach (var kfo in _kfoList)
                {
                    AddRowForKFO(kfo, sheetData);
                }
            }
        }

        /// <summary>
        ///  Creates and adds a row for the data in kfo
        /// </summary>
        /// <param name="kfo">KeywordFileOccurrence object representing the keyword found in the file</param>
        /// <param name="sheetData">SheetData object for the spreadsheet to be populated</param>
        private void AddRowForKFO(KeywordFileOccurrence kfo, SheetData sheetData)
        {
            Row newRow = MakeDataRow(kfo, _lastFilledRow + 1);
            sheetData.Append(newRow);
            _lastFilledRow += 1;
        }


        /// <summary>
        /// Creates and populates a row representing the column headers
        /// </summary>
        /// <param name="sheetData">SheetData object representing the spreadsheet to fill</param>
        /// <returns></returns>
        private Row MakeHeaderRow(SheetData sheetData)
        {
            Row row = new Row() { RowIndex = 1 };

            var columnHeaderCellIds = new List<string>() { "A1", "B1", "C1", "D1", "E1" };
            var headerTextList = new List<string>() { "Keyword", "File Name", "Slide Indices", "File Path", "Date Created" };

            for (int i = 0; i < columnHeaderCellIds.Count; i++)
            {
                Cell cell = MakeCell(headerTextList[i], columnHeaderCellIds[i], true);
                row.Append(cell);
            }

            _lastFilledRow = 1;

            return row;
        }

        /// <summary>
        /// Creates a row at index rowIndex in the spreadsheet and populates it with the data in kfo
        /// </summary>
        /// <param name="kfo">KeywordFileOccurrence object representing the occurrence of a keyword in a file</param>
        /// <param name="rowIndex">Excel row index (1-based) for the row</param>
        /// <returns></returns>
        private Row MakeDataRow(KeywordFileOccurrence kfo, uint rowIndex)
        {
            Row row = new Row() { RowIndex = rowIndex };
            var newCellList = new List<Cell>();

            newCellList.Add(MakeCell(kfo.Keyword, "A" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(Path.GetFileName(kfo.FilePath), "B" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(string.Join(",", kfo.SlideIndices), "C" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(kfo.FilePath, "D" + rowIndex.ToString(), false));
            newCellList.Add(MakeCell(kfo.CreationDate.ToString().Split()[0], "E" + rowIndex.ToString(), false));

            foreach (Cell newCell in newCellList)
            {
                row.Append(newCell);
            }

            return row;
        }

        /// <summary>
        /// Creates a cell object at cellIndex and fills it with the text specifiec in cellText
        /// </summary>
        /// <param name="cellText">text to enter into the cell</param>
        /// <param name="cellIndex">excel index for the cell (e.g. "C2")</param>
        /// <param name="bold">make text bold</param>
        /// <returns></returns>
        private Cell MakeCell(string cellText, string cellIndex, bool bold)
        {
            //create a new inline string cell
            Cell cell = new Cell() { CellReference = cellIndex };
            cell.DataType = CellValues.InlineString;

            //create a run for the bold text
            Run run1 = new Run();
            run1.Append(new Text(cellText));
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

        /// <summary>
        /// Creates an excel document in memory with a single worksheet. This document has no knowledge of the 
        /// storage location of the document at _outputPath until Spreadsheet.Close is implicitly called at the
        /// end of the using statement in WriteDictToFile().
        /// </summary>
        private void CreateDocument()
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
