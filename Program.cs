using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using WXML = DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Presentation;
using XMLyzeLibrary.Excel;
using XMLyzeLibrary.Word;

namespace XMLyze
{
    class XMLyze
    {
        static void Main(string[] args)
        {
            // Check command line arguments
            if (args.Length < 2)
            {
                Console.WriteLine("Usage: dotnet run original.xlsx new.docx");
                return;
            }

            // Paths
            string excelFilePath = $"docs/{args[0]}";
            string baseFileName = Path.GetFileNameWithoutExtension(args[0]);
            string wordFilePath = $"docs/{args[1]}";
            string imagesFolderPath = $"docs/{baseFileName}-imgs";

            // Open Excel file, create Word package
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(excelFilePath, false))
            using (WordprocessingDocument newPackage = WordprocessingDocument.Create(wordFilePath, WordprocessingDocumentType.Document))
            {
                if (spreadsheetDocument is null)
                    throw new ArgumentNullException(nameof(spreadsheetDocument));

                // Get excel parts
                WorkbookPart? workbookPart = spreadsheetDocument.WorkbookPart;
                if (workbookPart is null)
                    throw new ArgumentNullException(nameof(workbookPart));
                var sharedStringTable = workbookPart.SharedStringTablePart?.SharedStringTable;
                if (sharedStringTable is null)
                    throw new ArgumentNullException(nameof(sharedStringTable));
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Access the media folder and extract images
                List<string> imageFilePaths = EF.ExtractImages(excelFilePath, imagesFolderPath);

                // Populate new Word package
                (MainDocumentPart mainPart, WXML.Body body) = WF.PopulateNewWordPackage(newPackage, 1134, "blue");

                // Get excel data
                List<List<Cell>> rows = EF.GetRows(sheetData);
            }
        }
    }
}