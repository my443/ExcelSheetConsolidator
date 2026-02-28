using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator
{
    internal class ExcelOutput
    {

        private List<Cell> _mappings;
        private List<WorkbookExtraction> _outputData;
        public ExcelOutput(List<Cell> TemplateItems){ 
            _mappings = TemplateItems;
            _outputData = new List<WorkbookExtraction>();
        }

        public string[] GetListOfExcelFilesFromDirectory(string folderPath)
        {
            string[] excelFilesList = Directory.GetFiles(folderPath, "*.xlsx");
            return excelFilesList;
        }

        public void LoopThroughWorkbooks(string[] listOfFiles, string folderPath) {
            foreach (string file in listOfFiles) {
                ExtractDataFromWorkBook($@"{folderPath}\{file}");
            }
        }

        public void ExtractDataFromWorkBook(string workbookFilePath) {

            WorkbookExtraction workbookExtraction = new WorkbookExtraction();

            using (var workbook = new XLWorkbook(workbookFilePath))
            {
                foreach (Cell item in _mappings)
                {
                    var worksheet = workbook.Worksheet(item.SourceSheet);
                    var cell = worksheet.Cell(item.SourceReference);

                    // Get the value (as a generic object or specific type)
                    XLCellValue value = cell.Value;
                    item.CellValue = value;

                    workbookExtraction.Add(item);
                }                
            }

            _outputData.Add(workbookExtraction);
        }
    }
}
