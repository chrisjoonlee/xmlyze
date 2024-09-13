using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Linq;
using System.IO;
using System.IO.Compression;

namespace XMLyzeLibrary.Excel
{
    public static class EF
    {
        public static List<List<Cell>> GetRows(SheetData sheetData)
        {
            List<List<Cell>> fullRows = [];

            foreach (Row row in sheetData.Elements<Row>())
            {
                List<Cell> fullRow = [];
                int currentColumnIndex = 0;

                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (cell.CellReference != null)
                    {
                        string columnName = GetColumnName(cell.CellReference!);
                        int cellColumnIndex = GetColumnIndexFromName(columnName);

                        // Fill in missing cells
                        while (currentColumnIndex < cellColumnIndex)
                        {
                            fullRow.Add(new Cell());
                            currentColumnIndex++;
                        }

                        fullRow.Add(cell);
                        currentColumnIndex++;
                    }

                }

                // Ensure row has enough cells up to the desired column count
                // int totalColumns = 10; // Adjust based on your expected maximum column count
                // while (fullRow.Count < totalColumns)
                // {
                //     fullRow.Add(new Cell());
                // }

                fullRows.Add(fullRow);
            }

            return fullRows;
        }

        public static int GetColumnIndexFromName(string columnName)
        {
            int columnIndex = 0;
            foreach (char c in columnName)
            {
                columnIndex *= 26;
                columnIndex += (c - 'A' + 1);
            }
            return columnIndex - 1;
        }

        public static string GetColumnName(string cellReference)
        {
            // Extract column name from cell reference (e.g., "A1" -> "A")
            return new string(cellReference.Where(c => char.IsLetter(c)).ToArray());
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
                }
            }

            return imageFilePaths;
        }
    }
}