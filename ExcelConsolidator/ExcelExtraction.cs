using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelConsolidator
{
    internal class ExcelExtraction
    {
        private ExportTemplate _mappings;
        private ExportRowsCollection _outputData;
        public ExportRowsCollection ExportRowsCollection { get => _outputData; }
        public ExcelExtraction(ExportTemplate template)
        {
            _mappings = template;
            _outputData = new();
        }

        public ExportRowsCollection ExtractDataFromDirectory(string folderPath)
        {
            string[] filelist = GetListOfExcelFilesFromDirectory(folderPath);
            ExportRowsCollection outputData = LoopThroughWorkbooks(filelist, folderPath);

            return outputData;
        }

        private string[] GetListOfExcelFilesFromDirectory(string folderPath)
        {
            string[] excelFilesList = Directory.GetFiles(folderPath, "*.xlsx");
            return excelFilesList;
        }

        private ExportRowsCollection LoopThroughWorkbooks(string[] listOfFiles, string folderPath)
        {
            ExportRowsCollection outputData = new ExportRowsCollection();
            foreach (string file in listOfFiles)
            {
                var row = ExtractDataFromWorkBook($@"{file}");
                outputData.Add(row);
            }

            _outputData = outputData;
            return outputData;
        }

        private ExportRow ExtractDataFromWorkBook(string workbookFilePath)
        {

            ExportRow row = new ExportRow();

            using (var workbook = new XLWorkbook(workbookFilePath))
            {
                foreach (AbsoluteCell item in _mappings.TemplateItems)
                {
                    var worksheet = workbook.Worksheet(item.SourceSheet);
                    var cell = worksheet.Cell(item.SourceReference);

                    // Get the value (as a generic object or specific type)
                    XLCellValue value = cell.Value;
                    item.CellValue = value;

                    row.Add(item);
                }
            }

            return row;
        }
    }
}
