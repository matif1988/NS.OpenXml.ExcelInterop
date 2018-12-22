[![Latest version](https://img.shields.io/nuget/v/NS.OpenXml.ExcelInterop.svg)](https://www.nuget.org/packages/NS.OpenXml.ExcelInterop)
[![NuGet](https://img.shields.io/nuget/dt/NS.OpenXml.ExcelInterop.svg)](https://www.nuget.org/packages/NS.OpenXml.ExcelInterop)
[![master](https://img.shields.io/azure-devops/build/matif/Cronos/1/master.svg)](https://img.shields.io/azure-devops/build/matif/Cronos/1/master.svg)
[![MyGet](https://img.shields.io/azure-devops/release/matif/8e0bf57f-834e-410f-8211-93de0614324a/1/1.svg)](https://img.shields.io/azure-devops/release/matif/8e0bf57f-834e-410f-8211-93de0614324a/1/1.svg)
[![Build status](https://matif.visualstudio.com/Cronos/_apis/build/status/NS.OpenXml.ExcelInterop-CI)](https://matif.visualstudio.com/Cronos/_build/latest?definitionId=1)

# Open XML Excel Interop
The Open XML Excel Interop is a small .Net library that imports and exports excel file using Open XML SDK.

Table of Contents
-----------------

- [Dependencies](#dependencies)
- [Releases](#releases)
  - [Supported platforms](#supported-platforms)
- [Excel Import file](#excel-import-file)
- [Excel Export file](#excel-export-file)
- [If You Have Problems](#if-you-have-problems)
- [Support](#support)

Dependencies
------------
* [DocumentFormat.OpenXml](https://www.nuget.org/packages/DocumentFormat.OpenXml/)

Releases
--------------------------------

The official release NuGet packages for Open XML Excel Interop are [available on Nuget.org](https://www.nuget.org/packages/NS.OpenXml.ExcelInterop).

The NuGet package for the latest builds of the Open XML Excel Interop is available as a custom feed on MyGet. You can trust this package source, since the custom feed is locked and only this project feeds into the source. Stable releases here will be mirrored onto NuGet and will be identical.

Supported platforms
-----------------

This library supports many platforms. There are builds for .NET 4.5, .NET 4.6, and .NET Standard 2.0. The following platforms are currently supported:

|    Platform     | Minimum Version |
|-----------------|-----------------|
| .NET Framework  | 4.5             |
| .NET Core       | 1.0             |

Excel Import file
-----------------

Through the ExcelFileReader class, you can import your excel files. You have the possibility to recover the data in formats below:
* List of a dictionary composed of the name of the column and its value (with the possibility of including / excluding the header of the Excel table).
* DataTable that includes the columns and rows of the Excel table.

Excel Export file
-----------------

Through the ExcelFileWriter class, you can export your data to Excel file. This class includes a method called ExportToExcel that takes as input the parameters below:
* DataTable: Contains the data to export.
* WorkSheetName (string): Contains the work sheet name.
* WithFooter (bool): true if the last row is in bold, false otherwise.
This method returns a MemoryStream containing the excel file.

If You Have Problems
--------------------

If you want to report a problem (bug, behavior, build, distribution, feature request, etc...) with this class library built by this repository, please feel free to post a new issue and I will try to help.

Support
-------

This project is open source, it was developed to make the handling of Excel imports and exports more user-friendly. I thank you in advance for anyone in the community for the possible improvements of the solution as well as the report of possible bugs allowing me and any stakeholder to lead a continuous improvement of this product.
