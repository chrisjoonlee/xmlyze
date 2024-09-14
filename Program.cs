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
using XMLyzeLibrary.Interpreter;

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

    // Turn Excel data into code blocks
    List<IF.CodeBlock> codeBlocks = IF.GetCodeBlocksFromExcelFile(excelFilePath);

    // Populate new Word package
    (MainDocumentPart mainPart, WXML.Body body) = WF.PopulateNewWordPackage(newPackage, 1134, "blue");

    // Read through code blocks
    foreach (IF.CodeBlock codeBlock in codeBlocks)
    {
        Console.WriteLine(codeBlock);
        switch (codeBlock.Command)
        {
            case IF.Command.Paragraph:
                foreach (string text in codeBlock.Texts)
                    WF.AppendToBody(body, WF.Paragraph(text));
                break;
        }
    }
}
