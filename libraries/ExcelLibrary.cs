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
        public static string GetCellValue(Cell cell, WorkbookPart workbookPart)
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

        public static int GetColumnIndex(string cellReference)
        {
            string columnLetter = new string(cellReference.Where(char.IsLetter).ToArray());
            int columnIndex = 0;

            foreach (char letter in columnLetter)
            {
                columnIndex = columnIndex * 26 + (letter - 'A' + 1);
            }

            return columnIndex;
        }

        public static bool IsTextCell(Cell cell)
        {
            return cell.DataType != null && cell.DataType == CellValues.SharedString;
        }

        public static bool IsImageCell(Cell cell)
        {
            return cell.DataType != null
                && cell.DataType.Value == CellValues.Error
                && !string.IsNullOrEmpty(cell.GetAttribute("vm", "").Value);
        }

        public static bool IsNumberCell(Cell cell)
        {
            return cell != null && cell.CellValue != null && (cell.DataType == null || cell.DataType.Value == CellValues.Number);
        }

        public static List<string> ExtractImages(string excelFilePath, string imageFolderPath)
        {
            List<string> imageFilePaths = [];

            using (FileStream zipToOpen = new FileStream(excelFilePath, FileMode.Open))
            using (ZipArchive archive = new ZipArchive(zipToOpen, ZipArchiveMode.Read))
            {
                var mediaEntries = archive.Entries.Where(e => e.FullName.StartsWith("xl/media/")).ToList();

                if (!mediaEntries.Any())
                {
                    Console.WriteLine("No images found in the workbook.");
                }
                else
                {
                    foreach (var mediaEntry in mediaEntries)
                    {
                        using (Stream stream = mediaEntry.Open())
                        using (MemoryStream memoryStream = new MemoryStream())
                        {
                            stream.CopyTo(memoryStream);
                            byte[] imageBytes = memoryStream.ToArray();

                            // Create images folder
                            if (!Directory.Exists(imageFolderPath))
                                Directory.CreateDirectory(imageFolderPath);

                            string newImagePath = $"{imageFolderPath}/{mediaEntry.Name}";
                            imageFilePaths.Add(newImagePath);
                            File.WriteAllBytes(newImagePath, imageBytes);
                            Console.WriteLine($"Image saved to {newImagePath}");
                        }
                    }
                    Console.WriteLine();
                }
            }

            Console.WriteLine("IMAGE PATHS");
            foreach (string path in imageFilePaths)
                Console.WriteLine(path);
            Console.WriteLine();


            return imageFilePaths;
        }
    }
}