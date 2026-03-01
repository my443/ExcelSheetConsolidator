using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Services
{
    internal class ExtractionTemplate
    {
        public ExportTemplate TemplateItems { get; set; }
        public ExportTemplate GetTemplateItems(string filepath)
        {
            var items = new ExportTemplate();

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
                // This is just if it fails to open the workbook.
            }

            return items;
        }

        private ExportTemplate GetRowsData(IXLRangeRows rows)
        {
            var rowsData = new ExportTemplate();

            foreach (var row in rows)
            {
                rowsData.Add(MapRowToItem(row));
            }

            return rowsData;
        }

        private CellDefinition MapRowToItem(IXLRangeRow row)
        {
            CellDefinition newTemplateItem = new CellDefinition
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
