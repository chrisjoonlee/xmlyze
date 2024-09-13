﻿using DocumentFormat.OpenXml;
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
using XMLyzeLibrary.Interpreter;

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
            string excelFilePath = $"{args[0]}";
            string baseFileName = Path.GetFileNameWithoutExtension(args[0]);
            string wordFilePath = $"{args[1]}";
            string imagesFolderPath = $"{baseFileName}-imgs";

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

                // Get tokenized data
                List<List<string>> rows = EF.ReadExcelSheet(excelFilePath);
                List<List<Token>> tokenRows = [];
                foreach (List<string> row in rows)
                {
                    List<Token> tokens = EF.TokenizeRow(row);
                    tokenRows.Add(tokens);
                }

                // Read tokenized data
                foreach (List<Token> row in tokenRows)
                {
                    foreach (Token token in row)
                    {
                        switch (token.Type)
                        {
                            // Commands
                            case TokenType.Command:
                                if (IF.CommandDict.TryGetValue(token.Value, out IF.Command command))
                                {
                                    Console.WriteLine(command);
                                }
                                break;
                        }
                    }
                }
            }
        }
    }
}