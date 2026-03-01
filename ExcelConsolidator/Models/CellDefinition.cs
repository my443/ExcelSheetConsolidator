using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class CellDefinition
    {
        // A Definition of the Source, Destination, and value of each Cell
        // A cell is defined for an absolute location in a Workbook (Sheet and Reference) 
        //      Not just a reference (eg A1) on the current sheet.
        public string SourceSheet { get; set; }         // e.g. Sheet1, EmployeeData
        public string SourceReference { get; set; }     // e.g. A1, B27
        public string OutputSheet { get; set; }         // The Sheet where the new table is created. 
        public int OutputColumn { get; set; }           // Can only be a column number, because it will be the same each row. 
        public string OutputColumnName { get; set; }    // The name of the output column. 
        public XLCellValue CellValue { get; set; }      // The value object extracted from Excel
    }
}
