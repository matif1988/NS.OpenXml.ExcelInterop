/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NS.OpenXml.ExcelInterop.Helpers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NS.OpenXml.ExcelInterop.Managers
{
    /// <summary>
    /// The excel reader manager class
    /// </summary>
    internal class ExcelReaderManager : IDisposable
    {
        /// <summary>
        /// The spread sheet document
        /// </summary>
        protected SpreadsheetDocument _spreadsheetDocument = null;

        /// <summary>
        /// Loads the document.
        /// </summary>
        /// <param name="excelFilePathname">The excel file pathname.</param>
        /// <returns>true if document loading is sucessfull, false otherwise</returns>
        internal bool LoadDocument(string excelFilePathname)
        {
            _spreadsheetDocument = SpreadsheetDocument.Open(excelFilePathname, false);

            return _spreadsheetDocument != null;
        }

        /// <summary>
        /// Loads the document.
        /// </summary>
        /// <param name="excelFileStream">The excel file stream.</param>
        /// <returns>true if document loading is sucessfull, false otherwis</returns>
        internal bool LoadDocument(Stream excelFileStream)
        {
            _spreadsheetDocument = SpreadsheetDocument.Open(excelFileStream, false);

            return _spreadsheetDocument != null;
        }

        /// <summary>
        /// Gets the worksheet.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns>the selected Worksheet</returns>
        internal Worksheet GetWorksheet(string name)
        {
            if (_spreadsheetDocument == null)
                return null;

            WorksheetPart wspart = ExcelReaderHelper.GetWorksheetPartByName(_spreadsheetDocument, name);

            if (wspart == null)
                return null;

            return wspart.Worksheet;
        }

        /// <summary>
        /// Gets all worksheet.
        /// </summary>
        /// <returns>the selected Worksheet</returns>
        internal IEnumerable<Worksheet> GetAllWorksheet()
        {
            if (_spreadsheetDocument == null)
            {
                return null;
            }

            IEnumerable<WorksheetPart> worksSheetPartsList = ExcelReaderHelper.GetWorksheetParts(_spreadsheetDocument);
            if (worksSheetPartsList != null)
            {
                return worksSheetPartsList.Where(w => w?.Worksheet != null).Select(w => w.Worksheet);
            }

            return new List<Worksheet>();
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="column">The column.</param>
        /// <param name="row">The row.</param>
        /// <returns>object representing cell</returns>
        internal object GetCellValue(Worksheet worksheet, string column, uint row)
        {
            Cell cell = ExcelReaderHelper.GetCell(worksheet, column, row);
            if (cell == null)
            {
                return null;
            }

            return ExcelReaderHelper.GetCellValue(_spreadsheetDocument, cell);
        }

        /// <summary>
        /// Gets the cell value.
        /// </summary>
        /// <param name="cell">The cell.</param>
        /// <returns>The cell value as object.</returns>
        internal object GetCellValue(Cell cell)
        {
            return ExcelReaderHelper.GetCellValue(_spreadsheetDocument, cell);
        }

        /// <summary>
        /// Gets the row values.
        /// </summary>
        /// <param name="worksheet">The worksheet.</param>
        /// <param name="rowIndex">Index of the row.</param>
        /// <returns>A Dictionary containing row values</returns>
        internal Dictionary<string, object> GetRowValues(Worksheet worksheet, int rowIndex)
        {
            Dictionary<string, Cell> rowCells = ExcelReaderHelper.GetRowCells(worksheet, (uint)rowIndex);

            if (rowCells == null)
                return null;

            Dictionary<string, object> rowValues = new Dictionary<string, object>();

            foreach (KeyValuePair<string, Cell> element in rowCells)
                rowValues.Add(element.Key, ExcelReaderHelper.GetCellValue(_spreadsheetDocument, element.Value));

            return rowValues;
        }

        /// <summary>
        /// Gets work sheet rows.
        /// </summary>
        /// <param name="worksheet">The work sheet.</param>
        /// <returns>The work sheet rows.</returns>
        internal IEnumerable<Row> GetWorkSheetRows(Worksheet worksheet)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            return sheetData?.Descendants<Row>();
        }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            if (_spreadsheetDocument != null)
            {
                _spreadsheetDocument.Dispose();
            }
        }
    }
}
