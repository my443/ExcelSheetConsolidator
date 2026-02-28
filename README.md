# Excel Consolidator

## Purpose of this App

This utility (app) looks at directory of consitent Excel files and consolidates that data from each Workbook into one table. 

## Use Cases

* Consolidating a directory of budget templates from each department into one table. 
* Extracting data from Excel-based expense claims or petty cash submissions into a table for Accounts Payable to process. 
* Consolidating the monthly work-records from an entire office of consultants into one table that can be used for generating client's bills. 

## How to use this Utility

1. Create an Excel file which maps the data you want to extract and how it will appear in the export table. 
2. Put all of the files you want to conslidate into one directory. 
3. Use the command `ExcelConsolidator.exe "MappingTemplate=C:\My Mapping Template File.xlsx" "SourceFileDirectory=c:\Directory of source files" "OutputFilename=c:\OutputFilename.xlsx"`
The quotes are important in your command. You can put this command into a batch file so that you can reuse it whenever you want to. 

The ExcelConsolidator will read the MappingTemplate and consolidate all of the data from each Excel file in the Source File Directory 
into a single spreadsheet in the file that you have defined.