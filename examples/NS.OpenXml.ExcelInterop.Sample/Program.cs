using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace NS.OpenXml.ExcelInterop.Sample
{
    /// <summary>
    /// The program class
    /// </summary>
    class Program
    {
        /// <summary>
        /// The main method
        /// </summary>
        /// <param name="args">The arguments.</param>
        static void Main(string[] args)
        {
            // Retrieve project path
            string projectPath = AppDomain.CurrentDomain.BaseDirectory.Replace("\\bin\\Debug\\", string.Empty);

            // Retrieve template document.
            var templatePath = $"{projectPath}\\Template\\UserList.xlsx";
            var fileAsStream = new MemoryStream(File.ReadAllBytes(templatePath));

            // Import file with ExcelFileReader From File info to List of dictionary with header
            List<Dictionary<string, object>> dictionaryResultFromFileInfo = ExcelFileReader.ImportFileAsDictionary(new FileInfo(templatePath), "Users", true);

            // Import file with ExcelFileReader From File Stream to List of dictionary without header
            List<Dictionary<string, object>> dictionaryResultFromFileStream = ExcelFileReader.ImportFileAsDictionary(fileAsStream, "Users");

            // Import file with ExcelFileReader From File info to Data table
            DataTable dataTableResultFromFileInfo = ExcelFileReader.ImportFileAsDataTable(new FileInfo(templatePath), "Users");

            // Import file with ExcelFileReader From File Stream to Data table
            DataTable dataTableResultFromFileStream = ExcelFileReader.ImportFileAsDataTable(fileAsStream, "Users");

            // Export file with ExcelFileWriter
            string tempFilePath = $"{projectPath}\\App_Data\\{Guid.NewGuid()}.xlsx";
            var streamExport = ExcelFileWriter.ExportToExcel(dataTableResultFromFileStream, "Users");
            FileStream file = new FileStream(tempFilePath, FileMode.Create, FileAccess.Write);
            streamExport.WriteTo(file);
            file.Close();
            streamExport.Close();
        }
    }
}
