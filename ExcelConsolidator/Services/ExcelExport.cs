using ClosedXML.Excel;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Services
{
    internal class ExcelExport
    {
        public ExcelExport(string filePath, ExportRowsCollection rowsCollection) { 
            SaveDataToExcel(filePath, rowsCollection);
        }

        private void SaveDataToExcel(string filePath, ExportRowsCollection rowsCollection) {
            var distinctSheets = rowsCollection.Rows
                                    .SelectMany(row => row.Cells)     // Flatten all Cells from all Rows into one list
                                    .Select(cell => cell.OutputSheet) // Grab just the OutputSheet string
                                    .Distinct()                       // Remove duplicates
                                    .ToList();                        // Optional: Convert to List<string>

            using (var workbook = new XLWorkbook())
            {
                // Add each fo the worksheets
                foreach (var sheet in distinctSheets) {
                    var worksheet = workbook.Worksheets.Add(sheet);
                }


                // Writing to a specific cell (Row, Column) or address
                workbook.Worksheet("Sheet2").Cell("A2").Value = "Hello World!";
                

                // Save the file
                workbook.SaveAs(filePath);
            }
        }
    }
}
