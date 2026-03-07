using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using ExcelConsolidator.Models;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace ExcelConsolidator.Services
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
                foreach (CellDefinition item in _mappings.TemplateItems)
                {
                    try
                    {
                        var worksheet = workbook.Worksheet(item.SourceSheet);
                        var cell = worksheet.Cell(item.SourceReference);

                        var result = new CellDefinition
                        {
                            SourceSheet = item.SourceSheet,
                            SourceReference = item.SourceReference,
                            OutputSheet = item.OutputSheet,
                            OutputColumn = item.OutputColumn,
                            OutputColumnName = item.OutputColumnName,
                            CellValue = cell.Value
                        };

                        row.Add(result);
                    }
                    catch (Exception)
                    {
                        var errorResult = new CellDefinition
                        {
                            SourceSheet = item.SourceSheet,
                            SourceReference = item.SourceReference,
                            OutputSheet = item.OutputSheet,
                            OutputColumn = item.OutputColumn,
                            OutputColumnName = item.OutputColumnName,
                            CellValue = $"Error: {item.SourceSheet}!{item.SourceReference} not found in {workbookFilePath}."
                        };
                        row.Add(errorResult);
                    }
                }
            }

            return row;
        }
    }
}
