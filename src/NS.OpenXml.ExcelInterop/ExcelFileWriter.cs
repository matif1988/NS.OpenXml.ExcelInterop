/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NS.OpenXml.ExcelInterop.Managers;
using System;
using System.Data;
using System.IO;

namespace NS.OpenXml.ExcelInterop
{
    /// <summary>
    /// The excel file writer class
    /// </summary>
    public static class ExcelFileWriter
    {
        #region public methods

        /// <summary>
        /// Exports data table to excel file.
        /// </summary>
        /// <param name="datatable">The data table to export.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <param name="withFooter">true if the last row is footer and should be in font bold, false otherwise</param>
        /// <returns>The exported file as memory stream.</returns>
        public static MemoryStream ExportToExcel(DataTable datatable, string workSheetName, bool withFooter = false)
        {
            MemoryStream memoryStream = new MemoryStream();

            using (ExcelWriterManager excelReaderManager = new ExcelWriterManager())
            {
                if (!excelReaderManager.CreateDocument(memoryStream))
                {
                    throw new ArgumentException("Unable to create the excel file");
                }

                WorkbookPart workbookPart = excelReaderManager.CreateWorkBookPart();
                WorksheetPart worksheetPart = excelReaderManager.CreateWorksheetPart(workbookPart);
                excelReaderManager.CreateStylePart(workbookPart);
                Workbook workbook = excelReaderManager.CreateWorkbook(workbookPart.GetIdOfPart(worksheetPart), workSheetName);
                Worksheet worksheet = new Worksheet();

                // columns
                excelReaderManager.CreateWorksheetColumns(worksheet, datatable.Columns.Count);

                SheetData sheetData = new SheetData();

                // header
                excelReaderManager.CreateHeaderRow(sheetData, datatable);

                // content
                uint index = 2;
                int rowCount = withFooter ? datatable.Rows.Count - 1 : datatable.Rows.Count;
                excelReaderManager.CreateContent(sheetData, datatable, index, rowCount);

                // footer
                if (withFooter)
                {
                    excelReaderManager.CreateFooterRow(sheetData, index++, datatable.Rows[rowCount - 1]);
                }

                worksheet.Append(sheetData);
                worksheetPart.Worksheet = worksheet;
                worksheetPart.Worksheet.Save();

                excelReaderManager.AddWorkbookSaveAndClose(workbook);
            }

            memoryStream.Seek(0, SeekOrigin.Begin);
            return memoryStream;
        }

        #endregion
    }
}
