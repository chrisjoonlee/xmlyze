using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using System.IO.Compression;
using DocumentFormat.OpenXml.Wordprocessing;
using XMLyzeLibrary.Interpreter;

namespace XMLyzeLibrary.Excel
{
    // Receives a file path to an excel file
    // Reads through every row in the excel sheet
    // Returns rows of data as a list of lists of strings
    public static class EF
    {
        public static List<List<string>> ReadExcelSheet(string filePath)
        {
            var rowsData = new List<List<string>>();

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                if (document.WorkbookPart is null)
                    throw new Exception("Could not find Excel file");
                WorkbookPart workbookPart = document.WorkbookPart;

                if (workbookPart.Workbook.Sheets is null)
                    throw new Exception("Could not find Excel sheets");
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>() ?? throw new Exception("Could not find Excel sheet");
                string sheetId = sheet.Id?.Value ?? throw new Exception("Could not find Excel sheet id");

                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                foreach (Row row in sheetData.Elements<Row>())
                {
                    var rowData = new List<string>();
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int colIndex = GetColumnIndex(cell.CellReference!);
                        while (rowData.Count < colIndex - 1)
                        {
                            rowData.Add(string.Empty);
                        }
                        rowData.Add(GetCellValue(cell, workbookPart));
                    }
                    rowsData.Add(rowData);
                }
            }

            return rowsData;
        }

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (workbookPart.SharedStringTablePart is null)
                throw new Exception("Could not find shared string table in Excel sheet");

            string value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }

        private static int GetColumnIndex(string cellReference)
        {
            string columnLetter = new string(cellReference.Where(char.IsLetter).ToArray());
            int columnIndex = 0;

            foreach (char letter in columnLetter)
            {
                columnIndex = columnIndex * 26 + (letter - 'A' + 1);
            }

            return columnIndex;
        }

        // Receives a row of excel data
        // Turns the data into tokens
        // Returns a list of the tokens
        public static List<Token> TokenizeRow(List<string> row)
        {
            var tokens = new List<Token>();

            // Commands & arguments
            if (!string.IsNullOrWhiteSpace(row[0]))
            {
                // Commands
                if (!row[0].Trim().StartsWith("//"))
                    tokens.Add(new Token { Type = TokenType.Command, Value = row[0].Trim().ToLower() });

                // Arguments
                for (int i = 1; i < row.Count; i++)
                    if (!row[i].Trim().StartsWith("//") && !string.IsNullOrWhiteSpace(row[i]))
                        tokens.Add(new Token { Type = TokenType.Argument, Value = row[i].Trim() });
            }
            // Text
            else
                tokens.Add(new Token { Type = TokenType.Text, Value = row[1] });

            return tokens;
        }
    }
}