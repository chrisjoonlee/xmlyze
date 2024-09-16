using DocumentFormat.OpenXml.Packaging;
using XMLyzeLibrary.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;

namespace XMLyzeLibrary.Interpreter
{
    public static class IF
    {
        public enum TokenType
        {
            Command,
            Argument,
            Text,
            Image
        }

        public class Token
        {
            public TokenType Type { get; set; }
            public string Value { get; set; } = string.Empty;

            public override string ToString()
            {
                return $"Token(Type: {Type}, Value: \"{Value}\")";
            }

            public bool IsEmpty()
            {
                return string.IsNullOrEmpty(Value);
            }
        }

        // ADD NEW COMMANDS HERE
        public enum Command
        {
            Paragraph,
            Style
        }

        // ADD NEW COMMAND ALIASES HERE
        public static readonly Dictionary<string, Command> CommandDict = new()
        {
            { "paragraph", Command.Paragraph },
            { "p", Command.Paragraph },
            { "style", Command.Style },
            { "s", Command.Style }
        };

        // ADD NEW COMMAND ARGS HERE
        public static readonly Dictionary<Command, string[]> CommandArgsDict = new()
        {
            {Command.Paragraph, ["style"]},
            {Command.Style, ["name", "parent", "color", "size", "font"]}
        };


        public class CodeBlock
        {
            public Command Command { get; set; }
            public List<Argument> Arguments { get; set; } = [];
            public List<Token> Body { get; set; } = [];

            public void StripLeadingEmptyText()
            {
                while (Body.Count > 0 && Body[0].IsEmpty())
                    Body.RemoveAt(0);
            }

            public void StripTrailingEmptyText()
            {
                for (int i = Body.Count - 1; i >= 0; i--)
                {
                    if (Body[i].IsEmpty())
                        Body.RemoveAt(i);
                    else break;
                }
            }

            public override string ToString()
            {
                // Start with the command name
                var result = $"Command: {Command}\n";

                // Add arguments, if any
                if (Arguments.Count != 0)
                {
                    result += "Arguments:\n";
                    result += string.Join("\n", Arguments.Select(arg => $"  {arg}"));
                    result += "\n";
                }

                // Add text blocks, if any
                if (Body.Count != 0)
                {
                    result += "Body:\n";
                    result += string.Join("\n", Body.Select(item => $"  {item}"));
                }

                return result;
            }
        }

        public class Argument
        {
            public string Name { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;

            public override string ToString()
            {
                return $"{Name}: {Value}";
            }
        }

        public static List<CodeBlock> GetCodeBlocks(List<Token> tokens)
        {
            List<CodeBlock> codeBlocks = [];
            CodeBlock? currentCodeBlock = null;

            foreach (var token in tokens)
            {
                switch (token.Type)
                {
                    case TokenType.Command:
                        currentCodeBlock = new CodeBlock { Command = CommandDict[token.Value] };
                        codeBlocks.Add(currentCodeBlock);
                        break;

                    case TokenType.Argument:
                        if (currentCodeBlock == null)
                            throw new Exception("Argument not attached to a command");

                        // Get name and value of argument
                        string[] parts = token.Value.Split(['='], 2);
                        string name = parts[0].Trim().ToLower();
                        string value = parts[1].Trim();

                        // Check for unrecognized args
                        if (!CommandArgsDict[currentCodeBlock.Command].Contains(name))
                            throw new Exception($"{currentCodeBlock.Command} command does not have an argument called {name}");

                        currentCodeBlock?.Arguments.Add(new Argument { Name = name, Value = value });
                        break;

                    case TokenType.Text:
                        currentCodeBlock?.Body.Add(token);
                        break;
                }
            }

            foreach (CodeBlock codeBlock in codeBlocks)
            {
                if (codeBlock.Body.Count == 1) continue;

                // Strip leading and trailing empty text
                codeBlock.StripLeadingEmptyText();
                codeBlock.StripTrailingEmptyText();
            }

            return codeBlocks;
        }

        public static List<CodeBlock> GetCodeBlocksFromExcelFile(string excelFilePath)
        {
            List<Token> tokens = GetTokensFromExcelFile(excelFilePath);
            return GetCodeBlocks(tokens);
        }

        public static List<Token> GetTokensFromExcelFile(string filePath)
        {
            List<Token> tokens = [];

            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                // Validate file
                if (document.WorkbookPart is null)
                    throw new Exception("Could not find Excel file");
                WorkbookPart workbookPart = document.WorkbookPart;
                if (workbookPart.Workbook.Sheets is null)
                    throw new Exception("Could not find Excel sheets");
                Sheet sheet = workbookPart.Workbook.Sheets.GetFirstChild<Sheet>() ?? throw new Exception("Could not find Excel sheet");
                string sheetId = sheet.Id?.Value ?? throw new Exception("Could not find Excel sheet id");

                // Get sheet data
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheetId);
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                // Read sheet data
                foreach (Row row in sheetData.Elements<Row>())
                {
                    // int colCount = 0;

                    List<Cell> cells = row.Elements<Cell>().ToList();
                    bool isCommandRow = false;

                    for (int i = 0; i < cells.Count; i++)
                    {
                        Cell cell = cells[i];

                        // Create empty tokens for empty cells
                        // int colIndex = EF.GetColumnIndex(cell.CellReference!);
                        // while (colCount < colIndex - 1)
                        // {
                        //     tokens.Add(new Token
                        //     {
                        //         Type = TokenType.Text,
                        //         Value = ""
                        //     });
                        //     colCount++;
                        // }

                        string cellValue = EF.GetCellValue(cell, workbookPart);

                        // Comments
                        if (cellValue.Trim().StartsWith("//")) { }
                        // Empty space
                        else if (string.IsNullOrWhiteSpace(cellValue))
                        {
                            if (!isCommandRow)
                            {
                                tokens.Add(new Token
                                {
                                    Type = TokenType.Text,
                                    Value = ""
                                });
                            }

                            if (i == 0) isCommandRow = false;
                        }
                        else
                        {
                            // Commands
                            if (i == 0)
                            {
                                tokens.Add(new Token { Type = TokenType.Command, Value = cellValue.Trim().ToLower() });
                                isCommandRow = true;
                            }
                            // Arguments
                            else if (isCommandRow)
                            {
                                tokens.Add(new Token { Type = TokenType.Argument, Value = cellValue.Trim() });
                            }
                            else
                            {
                                // Images
                                if (EF.IsImageCell(cell))
                                {
                                    tokens.Add(new Token
                                    {
                                        Type = TokenType.Image,
                                        Value = EF.GetCellValue(cell, workbookPart)
                                    });
                                }
                                // Text
                                else
                                {
                                    tokens.Add(new Token
                                    {
                                        Type = TokenType.Text,
                                        Value = EF.GetCellValue(cell, workbookPart)
                                    });
                                }
                            }
                        }
                    }
                }
            }

            Console.WriteLine("TOKENS");
            foreach (Token token in tokens)
                Console.WriteLine(token);
            Console.WriteLine();

            return tokens;
        }
    }
}