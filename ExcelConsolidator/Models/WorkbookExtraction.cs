using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class WorkbookExtraction
    {
        // A list of cells that are extracted from a single workbook
        // Each cell has its own Worksheet defined,
        //      so this could be rows on multiple worksheets.
        List<Cell> Cells {  get; set; }

        public void Add(Cell cell) { 
            Cells.Add(cell);
        }
    }
}
