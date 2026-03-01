using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class ExportTemplate
    {
        public List<AbsoluteCell> TemplateItems = new();

        public void Add(AbsoluteCell cell) {
            TemplateItems.Add(cell);
        }
    }
}
