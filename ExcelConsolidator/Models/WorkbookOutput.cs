using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class WorkbookOutput
    {
        // A list of Workbook Extractions that will be tranformed into rows in the output workbook. 
        List<WorkbookExtraction> Rows { get; set; }
    }
}
