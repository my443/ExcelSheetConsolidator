using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Services
{
    internal class ExcelExport
    {
        public ExcelExport(string filePath, ExportRowsCollection rowsCollection, ExportTemplate template) { 
            SaveDataToExcel(filePath, rowsCollection, template);
        }

        private void SaveDataToExcel(string filePath, ExportRowsCollection rowsCollection, ExportTemplate template) {
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

                // Add the column headings for each one.
                foreach (var heading in template.TemplateItems) {
                    workbook.Worksheet(heading.OutputSheet).Cell(1, heading.OutputColumn).Value = heading.OutputColumnName;
                }

                int rowNumber = 2;

                foreach (var row in rowsCollection.Rows) {
                    foreach (var cell in row.Cells) {
                        workbook.Worksheet(cell.OutputSheet).Cell(rowNumber, cell.OutputColumn).Value = cell.CellValue;
                    }
                    rowNumber++;
                }               

                // Save the file
                workbook.SaveAs(filePath);
            }
        }
    }
}
