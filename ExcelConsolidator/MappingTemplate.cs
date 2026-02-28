using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator
{
    internal class MappingTemplate
    {
        public List<Models.Cell> TemplateItems { get; set; }
        public List<Models.Cell> GetTemplateItems(string filepath)
        {
            var items = new List<Models.Cell>();

            try
            {
                using (var workbook = new XLWorkbook(filepath))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(1);
                    IXLRangeRows rows = worksheet.Range(2, 1, worksheet.LastRowUsed().RowNumber(), 5).Rows(); // Starts at row 2 (Skip header)

                    items = GetRowsData(rows);
                }
            }
            catch {
                // This is just if it fails to open the book.
            }

            return items;
        }

        private List<Models.Cell> GetRowsData(IXLRangeRows rows)
        {
            var rowsData = new List<Models.Cell>();

            foreach (var row in rows)
            {
                // If the first cell is empty, we treat this as a "blank row" and stop
                // This isn't needed because we used worksheet.LastRowUsed() above.
                //if (row.Cell(1).IsEmpty()) break;

                rowsData.Add(MapRowToItem(row));
            }

            return rowsData;
        }

        private Models.Cell MapRowToItem(IXLRangeRow row)
        {
            Models.Cell newTemplateItem = new Models.Cell
            {
                SourceSheet = row.Cell(1).GetValue<string>(),
                SourceReference = row.Cell(2).GetValue<string>(),
                OutputSheet = row.Cell(3).GetValue<string>(),
                OutputColumn = row.Cell(4).GetValue<int>(),
                OutputColumnName = row.Cell(5).GetValue<string>()
            };

            return newTemplateItem;
        }
    }
}
