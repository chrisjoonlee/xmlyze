using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.IO;
using System.IO.Compression;
using XMLyzeLibrary.Excel;

namespace XMLyzeLibrary.Interpreter
{
    public static class IF
    {
        public enum TokenType
        {
            Command,
            Argument,
            Text
        }

        public class Token
        {
            public TokenType Type { get; set; }
            public string Value { get; set; } = string.Empty;

            public override string ToString()
            {
                return $"Token(Type: {Type}, Value: \"{Value}\")";
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
            {"style", Command.Style},
            {"s", Command.Style}
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
            public List<string> Texts { get; set; } = [];

            public void StripLeadingEmptyText()
            {
                while (Texts.Count > 0 && string.IsNullOrEmpty(Texts[0]))
                    Texts.RemoveAt(0);
            }

            public void StripTrailingEmptyText()
            {
                for (int i = Texts.Count - 1; i >= 0; i--)
                {
                    if (string.IsNullOrEmpty(Texts[i]))
                        Texts.RemoveAt(i);
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
                if (Texts.Count != 0)
                {
                    result += "Text:\n";
                    result += string.Join("\n", Texts.Select(text => $"  {text}"));
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
            {
                if (!row[1].Trim().StartsWith("//"))
                    tokens.Add(new Token { Type = TokenType.Text, Value = row[1] });
            }

            return tokens;
        }

        public static List<Token> GetTokens(List<List<string>> rows)
        {
            List<Token> tokens = [];
            foreach (List<string> row in rows)
                tokens.AddRange(IF.TokenizeRow(row));
            foreach (Token token in tokens)
                Console.WriteLine(token);
            return tokens;
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
                        currentCodeBlock?.Texts.Add(token.Value);
                        break;
                }
            }

            foreach (CodeBlock codeBlock in codeBlocks)
            {
                if (codeBlock.Texts.Count == 1) continue;

                // Strip leading and trailing empty text
                codeBlock.StripLeadingEmptyText();
                codeBlock.StripTrailingEmptyText();
            }

            return codeBlocks;
        }

        public static List<CodeBlock> GetCodeBlocksFromExcelFile(string excelFilePath)
        {
            List<List<string>> rows = EF.ReadExcelSheet(excelFilePath);
            List<Token> tokens = GetTokens(rows);
            return GetCodeBlocks(tokens);
        }
    }
}