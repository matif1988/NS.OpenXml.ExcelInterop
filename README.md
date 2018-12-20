[![Latest version](https://img.shields.io/nuget/v/NS.OpenXml.ExcelInterop.svg)](https://www.nuget.org/packages/NS.OpenXml.ExcelInterop)
[![NuGet](https://img.shields.io/nuget/dt/NS.OpenXml.ExcelInterop.svg)](https://www.nuget.org/packages/NS.OpenXml.ExcelInterop)
[![master](https://img.shields.io/azure-devops/build/matif/Cronos/1/master.svg)](https://img.shields.io/azure-devops/build/matif/Cronos/1/master.svg)
[![MyGet](https://img.shields.io/azure-devops/release/matif/8e0bf57f-834e-410f-8211-93de0614324a/1/1.svg)](https://img.shields.io/azure-devops/release/matif/8e0bf57f-834e-410f-8211-93de0614324a/1/1.svg)

# What is NS.OpenXml.ExcelInterop ?
NS.OpenXml.ExcelInterop is a small .Net library that imports and exports excel file using open xml.

This library supports both **.Net Framework 4.6**.

Depends on : 
* [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

# Excel Import file
Through the ExcelFileReader class, you can import your excel files. You have the possibility to recover the data in formats below:
* List of a dictionary composed of the name of the column and its value (with the possibility of including / excluding the header of the Excel table).
* DataTable that includes the columns and rows of the Excel table.

# Excel Export file
Through the ExcelFileWriter class, you can export your data to Excel file. This class includes a method called ExportToExcel that takes as input the parameters below:
* DataTable: Contains the data to export.
* WorkSheetName (string): Contains the work sheet name.
* WithFooter (bool): true if the last row is in bold, false otherwise.
This method returns a MemoryStream containing the excel file.

# Support
This project is open source, it was developed to make the handling of Excel imports and exports more user-friendly. I thank you in advance for anyone in the community for the possible improvements of the solution as well as the report of possible bugs allowing me and any stakeholder to lead a continuous improvement of this product.
