using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class ExportRow
    {
        // A list of cells that are extracted from a single workbook
        // Each cell has its own Worksheet defined,
        //      so this could be rows on multiple worksheets.
        List<AbsoluteCell> Cells { get; set; } = new();

        public void Add(AbsoluteCell cell) { 
            Cells.Add(cell);
        }
    }
}
