using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class ExportRowsCollection
    {
        // A list of Workbook Extractions that will be tranformed into rows in the output workbook. 
        List<ExportRow> Rows { get; set; } = new();

        public void Add(ExportRow row)
        {
            Rows.Add(row);
        }
    }
}
