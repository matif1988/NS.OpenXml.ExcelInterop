/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NS.OpenXml.ExcelInterop.Helpers;
using System;
using System.Data;
using System.IO;

namespace NS.OpenXml.ExcelInterop.Managers
{
    /// <summary>
    /// The excel writer manager class
    /// </summary>
    internal class ExcelWriterManager : IDisposable
    {
        /// <summary>
        /// The spread sheet document
        /// </summary>
        protected SpreadsheetDocument _spreadsheetDocument = null;

        /// <summary>
        /// Creates the document.
        /// </summary>
        /// <param name="excelFileStream">The excel file stream.</param>
        /// <returns>true if document is sucessfully created, false otherwise</returns>
        internal bool CreateDocument(Stream excelFileStream)
        {
            _spreadsheetDocument = SpreadsheetDocument.Create(excelFileStream, SpreadsheetDocumentType.Workbook);

            return _spreadsheetDocument != null;
        }

        /// <summary>
        /// Creates work book part
        /// </summary>
        /// <returns>The created workbookpart.</returns>
        internal WorkbookPart CreateWorkBookPart()
        {
            return _spreadsheetDocument.AddWorkbookPart();
        }

        /// <summary>
        /// Add workbook, save and close spread sheet document
        /// </summary>
        /// <param name="workbook">The work book.</param>
        /// <returns>true if document is sucessfully saved, false otherwise</returns>
        internal bool AddWorkbookSaveAndClose(Workbook workbook)
        {
            if (workbook != null)
            {
                _spreadsheetDocument.WorkbookPart.Workbook = workbook;
                _spreadsheetDocument.WorkbookPart.Workbook.Save();
                _spreadsheetDocument.Close();
                return true;
            }

            return false;
        }

        /// <summary>
        /// Creates work sheet part.
        /// </summary>
        /// <param name="workbookPart">The work book part.</param>
        /// <returns>The work sheet part.</returns>
        internal WorksheetPart CreateWorksheetPart(WorkbookPart workbookPart)
        {
            return workbookPart.AddNewPart<WorksheetPart>();
        }

        /// <summary>
        /// Creates style part
        /// </summary>
        /// <param name="workbookPart">The work book part.</param>
        internal void CreateStylePart(WorkbookPart workbookPart)
        {
            WorkbookStylesPart stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet styles = ExcelWriterHelper.CreateStylesheet();
            styles.Save(stylesPart);
        }

        /// <summary>
        /// Creates work book.
        /// </summary>
        /// <param name="relId">The relation identifier.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <returns>The created workbook.</returns>
        internal Workbook CreateWorkbook(string relId, string workSheetName)
        {
            Workbook workbook = new Workbook();
            FileVersion fileVersion = new FileVersion { ApplicationName = "Microsoft Office Excel" };
            Sheets sheets = new Sheets();
            Sheet sheet = new Sheet { Name = workSheetName, SheetId = 1, Id = relId };
            sheets.Append(sheet);
            workbook.Append(fileVersion);
            workbook.Append(sheets);

            return workbook;
        }

        /// <summary>
        /// Creates work sheet columns.
        /// </summary>
        /// <param name="worksheet">The work sheet.</param>
        /// <param name="numCols">The count of columns.</param>
        internal void CreateWorksheetColumns(Worksheet worksheet, int numCols)
        {
            Columns columns = new Columns();
            for (int col = 0; col < numCols; col++)
            {
                Column column = ExcelWriterHelper.CreateColumnData((uint)col + 1, (uint)numCols + 1, 25);
                columns.Append(column);
            }

            worksheet.Append(columns);
        }

        /// <summary>
        /// Creates header row.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="datatable">The data table.</param>
        internal void CreateHeaderRow(SheetData sheetData, DataTable datatable)
        {
            Row headerRow = ExcelWriterHelper.CreateHeaderRow(1, datatable);
            sheetData.AppendChild(headerRow);
        }

        /// <summary>
        /// Creates Content rows.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="datatable">The data table.</param>
        /// <param name="index">The index.</param>
        /// <param name="rowCount">The row count.</param>
        internal void CreateContent(SheetData sheetData, DataTable datatable, uint index, int rowCount)
        {
            for (int i = 0; i < rowCount; i++)
            {
                Row contentRow = ExcelWriterHelper.CreateContentRow(index++, datatable.Rows[i]);
                sheetData.AppendChild(contentRow);
            }
        }

        /// <summary>
        /// Creates footer row.
        /// </summary>
        /// <param name="sheetData">The sheet data.</param>
        /// <param name="index">The index.</param>
        /// <param name="footer">The footer.</param>
        internal void CreateFooterRow(SheetData sheetData, uint index, DataRow footer)
        {
            Row footerRow = ExcelWriterHelper.CreateFooterRow(index, footer);
            sheetData.AppendChild(footerRow);
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
