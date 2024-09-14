using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using D = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using DP = DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using System.Linq;
using System.IO;
using System.IO.Compression;

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

        public enum Command
        {
            Paragraph
        }

        public static readonly Dictionary<string, Command> CommandDict = new Dictionary<string, Command>
        {
            { "paragraph", Command.Paragraph },
            { "p", Command.Paragraph }
        };

        // Receives a row of excel data
        // Turns the data into tokens
        // Returns a list of the tokens


        public class CodeBlock
        {
            public Command Command { get; set; }
            public List<string> Arguments { get; set; } = [];
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
                    result += string.Join("\n", Arguments.Select(arg => $"  - {arg}"));
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
                        currentCodeBlock?.Arguments.Add(token.Value);
                        break;

                    case TokenType.Text:
                        currentCodeBlock?.Texts.Add(token.Value);
                        break;
                }
            }

            foreach (CodeBlock codeBlock in codeBlocks)
            {
                // Strip leading and trailing empty text
                codeBlock.StripLeadingEmptyText();
                codeBlock.StripTrailingEmptyText();
            }

            return codeBlocks;
        }
    }
}