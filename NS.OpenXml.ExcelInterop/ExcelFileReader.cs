/* Copyright (C) Mohammed ATIF https://github.com/matif1988/ns.openxml.excelInterop - All Rights Reserved */
using DocumentFormat.OpenXml.Spreadsheet;
using NS.OpenXml.ExcelInterop.Managers;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace NS.OpenXml.ExcelInterop
{
    /// <summary>
    /// The open xml excel file reader
    /// </summary>
    public static class ExcelFileReader
    {
        #region public methods

        /// <summary>
        /// Loads data from excel file
        /// </summary>
        /// <param name="file">The file info.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <param name="includeHeader">The include header.</param>
        /// <returns>The file data as list of dictionary.</returns>
        public static List<Dictionary<string, object>> ImportFileAsDictionary(FileInfo file, string workSheetName, bool includeHeader = false)
        {
            List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();

            using (ExcelReaderManager excelReaderManager = new ExcelReaderManager())
            {
                if (!excelReaderManager.LoadDocument(file.FullName))
                {
                    throw new ArgumentException("Unable to open the file");
                }

                dataList = ImportFileAsDictionary(excelReaderManager, workSheetName, includeHeader);
            }

            return dataList;
        }

        /// <summary>
        /// Loads data from excel file
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <param name="includeHeader">The include header.</param>
        /// <returns>The file data as list of dictionary.</returns>
        public static List<Dictionary<string, object>> ImportFileAsDictionary(Stream fileStream, string workSheetName, bool includeHeader = false)
        {
            List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();

            using (ExcelReaderManager excelReaderManager = new ExcelReaderManager())
            {
                if (!excelReaderManager.LoadDocument(fileStream))
                {
                    throw new ArgumentException("Unable to open the file");
                }

                dataList = ImportFileAsDictionary(excelReaderManager, workSheetName, includeHeader);
            }

            return dataList;
        }

        /// <summary>
        /// Imports data from excel file
        /// </summary>
        /// <param name="file">The file info.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <returns>The file data as data table.</returns>
        public static DataTable ImportFileAsDataTable(FileInfo file, string workSheetName)
        {
            DataTable dataTable = new DataTable();

            using (ExcelReaderManager excelReaderManager = new ExcelReaderManager())
            {
                if (!excelReaderManager.LoadDocument(file.FullName))
                {
                    throw new ArgumentException("Unable to open the file");
                }

                dataTable = ImportFileAsDataTable(excelReaderManager, workSheetName);
            }

            return dataTable;
        }

        /// <summary>
        /// Imports data from excel file
        /// </summary>
        /// <param name="fileStream">The file stream.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <returns>The file data as data table.</returns>
        public static DataTable ImportFileAsDataTable(Stream fileStream, string workSheetName)
        {
            DataTable dataTable = new DataTable();

            using (ExcelReaderManager excelReaderManager = new ExcelReaderManager())
            {
                if (!excelReaderManager.LoadDocument(fileStream))
                {
                    throw new ArgumentException("Unable to open the file");
                }

                dataTable = ImportFileAsDataTable(excelReaderManager, workSheetName);
            }

            return dataTable;
        }

        #endregion

        #region private methods

        /// <summary>
        /// Imports file as dictionary
        /// </summary>
        /// <param name="excelReaderManager">The excel reader manager.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <param name="includeHeader">The include header.</param>
        /// <returns>The file data as list of dictionary.</returns>
        static List<Dictionary<string, object>> ImportFileAsDictionary(ExcelReaderManager excelReaderManager, string workSheetName, bool includeHeader = false)
        {
            List<Dictionary<string, object>> dataList = new List<Dictionary<string, object>>();

            Worksheet workSheet = excelReaderManager.GetWorksheet(workSheetName);
            if (workSheet == null)
            {
                throw new ArgumentException("The sheet name can not be found");
            }

            IEnumerable<Row> rows = excelReaderManager.GetWorkSheetRows(workSheet);
            if (rows?.Any() == true)
            {
                int minRow = includeHeader ? 1 : 2;
                for (int iRow = minRow; iRow < rows.Count() + 1; iRow++)
                {
                    var rowValue = excelReaderManager.GetRowValues(workSheet, iRow);
                    if (rowValue != null)
                        dataList.Add(rowValue);
                }
            }

            return dataList;
        }

        /// <summary>
        /// Imports file as data table.
        /// </summary>
        /// <param name="excelReaderManager">The excel reader manager.</param>
        /// <param name="workSheetName">The work sheet name.</param>
        /// <returns>The file data as data table.</returns>
        static DataTable ImportFileAsDataTable(ExcelReaderManager excelReaderManager, string workSheetName)
        {
            DataTable dataTable = new DataTable();

            Worksheet workSheet = excelReaderManager.GetWorksheet(workSheetName);
            if (workSheet == null)
            {
                throw new ArgumentException("The sheet name can not be found");
            }

            IEnumerable<Row> rows = excelReaderManager.GetWorkSheetRows(workSheet);
            if (rows?.Any() == true)
            {
                // Retrieve file columns
                foreach (Cell cell in rows.ElementAt(0))
                {
                    var cellValue = excelReaderManager.GetCellValue(cell);
                    if (cellValue != null)
                        dataTable.Columns.Add(cellValue.ToString());
                }

                // Retrieve file rows
                foreach (Row row in rows.Skip(1)) //this will also include your header row...
                {
                    DataRow tempRow = dataTable.NewRow();

                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        tempRow[i] = excelReaderManager.GetCellValue(row.Descendants<Cell>().ElementAt(i));
                    }

                    dataTable.Rows.Add(tempRow);
                }
            }

            return dataTable;
        }

        #endregion
    }
}
