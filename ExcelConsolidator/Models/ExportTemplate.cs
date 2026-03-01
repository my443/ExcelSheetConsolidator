using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator.Models
{
    internal class ExportTemplate
    {
        public List<CellDefinition> TemplateItems = new();

        public void Add(CellDefinition cell) {
            TemplateItems.Add(cell);
        }
    }
}
