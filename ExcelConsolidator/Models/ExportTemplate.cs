using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class ExportTemplate
    {
        List<AbsoluteCell> TemplateItems = new();

        public void Add(AbsoluteCell cell) {
            TemplateItems.Add(cell);
        }
    }
}
