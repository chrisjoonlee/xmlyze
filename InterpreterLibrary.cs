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

    public static class IF
    {
        public enum Command
        {
            Paragraph
        }

        public static readonly Dictionary<string, Command> CommandDict = new Dictionary<string, Command>
        {
            { "paragraph", Command.Paragraph },
            { "p", Command.Paragraph }
        };
    }
}