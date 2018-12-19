/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace NS.OpenXml.ExcelInterop.Helpers
{
    /// <summary>
    /// The excel reader helper that allows to reads excel file in Open Xml format 
    /// </summary>
    internal static class ExcelReaderHelper
    {
        /// <summary>
        /// Max length of a worksheet name tab.
        /// </summary>
        private const int MaxWorksheetNameLength = 31;

        /// <summary>
        /// Checks that a cell is not null
        /// </summary>
        /// <param name="cell">excel sheet cell</param>
        /// <returns>true if not null otherwise false</returns>
        internal static bool CheckCell(Cell cell)
        {
            if (cell != null) return true;

            return false;
        }

        /// <summary>
        /// extracts worksheet parts from worksheet
        /// </summary>
        /// <param name="doc">worksheet</param>
        /// <returns>list of worksheet part</returns>
        internal static IEnumerable<WorksheetPart> GetWorksheetParts(SpreadsheetDocument doc)
        {
            foreach (var elem in doc.WorkbookPart.Workbook.Sheets)
            {
                Sheet sheet = elem as Sheet;
                if (sheet != null) yield return (WorksheetPart)doc.WorkbookPart.GetPartById(sheet.Id);
            }
        }

        /// <summary>
        /// Extracts row index from cell name
        /// </summary>
        /// <param name="cellName">cell name</param>
        /// <returns>value of row index</returns>
        internal static uint GetRowIndex(string cellName)
        {
            // Create a regular expression to match the row index portion the cell name.
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        /// <summary>
        /// Gets the index of the column.
        /// </summary>
        /// <param name="cellName">Name of the cell.</param>
        /// <returns>index of the column</returns>
        internal static uint GetColumnIndex(string cellName)
        {
            string columnName = GetColumnName(cellName);

            uint index = 0U;
            string columnNameUpper = columnName.ToUpper(CultureInfo.InvariantCulture);
            int length = columnNameUpper.Length;

            for (int i = 0; i < length; i++)
            {
                index += (columnNameUpper[length - i - 1] - 64U) * (uint)Math.Pow(26, i);
            }

            return index;
        }

        /// <summary>
        /// Get Column Name from column index
        /// </summary>
        /// <param name="columnIndex">Column index</param>
        /// <returns>Name of the column</returns>
        internal static string GetColumnName(int columnIndex)
        {
            int x = columnIndex / 26;
            string columnName;

            if (x > 0)
            {
                columnName = GetColumnName(x - 1) + (char)((columnIndex % 26) + 65);
            }
            else
            {
                columnName = string.Empty + (char)((columnIndex % 26) + 65);
            }

            return columnName;
        }

        /// <summary>
        /// extracts column name from cell name
        /// </summary>
        /// <param name="cellName">cell name</param>
        /// <returns>column index A ..Z</returns>
        internal static string GetColumnName(string cellName)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        /// <summary>
        /// select a worksheet cell with column and row index
        /// </summary>
        /// <param name="workSheet">worksheet</param>
        /// <param name="columnName">column index</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>cell object matching these indexes</returns>
        internal static Cell GetCell(Worksheet workSheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(workSheet, rowIndex);

            if (row == null) return null;

            return row.Elements<Cell>().Where(c => string.Compare(c.CellReference.Value, columnName + rowIndex, true) == 0).FirstOrDefault();
        }

        /// <summary>
        /// extract value of cell in the right type
        /// </summary>
        /// <param name="doc">Spread sheet Document</param>
        /// <param name="cell">cell to evaluate</param>
        /// <returns>cell value</returns>
        internal static object GetCellValue(SpreadsheetDocument doc, Cell cell)
        {
            if (cell.DataType == null)
            {
                if (cell.CellValue != null)
                {
                    return cell.CellValue.Text;
                }

                return null;
            }

            if (cell.DataType == CellValues.Boolean)
            {
                bool b = false;

                if (cell.CellValue != null && cell.CellValue.Text == Constant.BooleanTrueValue)
                    b = true;

                return b;
            }

            if (cell.DataType == CellValues.Date)
            {
                throw new NotImplementedException("Date Cell Type not implemented");
            }

            if (cell.DataType == CellValues.InlineString)
            {
                if (cell.CellValue != null)
                    return cell.CellValue.Text;

                if (cell.InlineString != null)
                    return cell.InlineString.Text.InnerText;

                return null;
            }

            if (cell.DataType == CellValues.Number)
            {
                throw new NotImplementedException("Number Cell Type not implemented");
            }

            if (cell.DataType == CellValues.SharedString)
            {
                var stringTable = doc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                string str = null;

                if (stringTable != null)
                {
                    str = stringTable.SharedStringTable.ElementAt(int.Parse(cell.CellValue.Text)).InnerText;
                }

                if (str == null)
                {
                    throw new ArgumentException("The Excel file is corrupted (can't find the Shared String in the string table)");
                }

                return str;
            }

            if (cell.DataType == CellValues.String)
            {
                if (cell.CellValue != null)
                    return cell.CellValue.Text;

                return null;
            }

            return null;
        }

        /// <summary>
        /// Given a worksheet and a row index, return the row.
        /// </summary>
        /// <param name="workSheet">worksheet </param>
        /// <param name="rowIndex">row index</param>
        /// <returns>targeted row object</returns>
        internal static Row GetRow(Worksheet workSheet, uint rowIndex)
        {
            return workSheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        /// <summary>
        /// Given a worksheet and a row index, return true if data exists after row index.
        /// </summary>
        /// <param name="workSheet">worksheet </param>
        /// <param name="rowIndex">row index</param>
        /// <returns>True if data exists after row index</returns>
        internal static bool DataExistsAfterRowIndex(Worksheet workSheet, uint rowIndex)
        {
            IEnumerable<Row> rows = workSheet.GetFirstChild<SheetData>().Elements<Row>().Where(r => r.RowIndex > rowIndex);
            if (!rows.Any())
                return false;

            int cellsWithData = (from r in rows
                                 from c in r.Elements<Cell>()
                                 where c.CellValue != null
                                 select c).Count();

            return cellsWithData > 0;
        }

        /// <summary>
        /// select a worksheet cells with a given row index
        /// </summary>
        /// <param name="workSheet">worksheet</param>
        /// <param name="rowIndex">row index</param>
        /// <returns>a dictionary where the key is the column name</returns>
        internal static Dictionary<string, Cell> GetRowCells(Worksheet workSheet, uint rowIndex)
        {
            Row row = GetRow(workSheet, rowIndex);

            if (row == null) return null;

            return row.Elements<Cell>().ToDictionary(c => GetColumnName(c.CellReference));
        }

        /// <summary>
        /// Get sheet by name
        /// </summary>
        /// <param name="document">document</param>
        /// <param name="sheetName">name</param>
        /// <returns>sheet</returns>
        internal static Sheet GetSheetByName(SpreadsheetDocument document, string sheetName)
        {
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (!sheets.Any())
            {
                // The specified worksheet does not exist.
                return null;
            }

            return sheets.First();
        }

        /// <summary>
        /// Get the worksheetPart by Name
        /// </summary>
        /// <param name="document">document to check</param>
        /// <param name="sheetName">sheet name to find</param>
        /// <returns>found WorksheetPart</returns>
        internal static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            Sheet sheet = GetSheetByName(document, sheetName);
            if (sheet == null)
                return null;

            return (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
        }

        /// <summary>
        /// Get the worksheetPart by Name
        /// </summary>
        /// <param name="document">document to check</param>
        /// <param name="sheet">sheet concerned</param>
        /// <returns>found WorksheetPart</returns>
        internal static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, Sheet sheet)
        {
            return (WorksheetPart)document.WorkbookPart.GetPartById(sheet.Id.Value);
        }

        /// <summary>
        /// Inserts a new worksheet into a workbook
        /// </summary>
        /// <param name="workbookPart">workbook part</param>
        /// <param name="sheetName">sheet name</param>
        /// <returns>Inserted sheet</returns>
        internal static Sheet InsertWorksheet(WorkbookPart workbookPart, string sheetName)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            return sheet;
        }

        /// <summary>
        /// Gets a valid name for a sheet.
        /// </summary>
        /// <param name="proposedName">Proposed name.</param>
        /// <returns>The valid name.</returns>
        internal static string GetValidSheetName(string proposedName)
        {
            if (string.IsNullOrWhiteSpace(proposedName))
            {
                return string.Empty;
            }

            string name = proposedName.Replace('/', '|');

            if (name.Length > MaxWorksheetNameLength)
            {
                return name.Substring(0, 31);
            }

            return name;
        }

        /// <summary>
        /// Duplicates the worksheet.
        /// </summary>
        /// <param name="workbookPart">The workbook part.</param>
        /// <param name="srcWorksheetPart">The worksheet part.</param>
        /// <param name="sheetName">Name of the sheet.</param>
        /// <returns>Dupplicated sheet</returns>
        internal static Sheet DuplicateWorksheet(WorkbookPart workbookPart, WorksheetPart srcWorksheetPart, string sheetName)
        {
            // Add a blank WorksheetPart.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = (Worksheet)srcWorksheetPart.Worksheet.CloneNode(true);

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Any())
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);

            return sheet;
        }
    }
}
