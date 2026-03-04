using System;
using System.IO;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Examples
{
    /// <summary>
    /// Utility to inspect Excel file headers for debugging
    /// </summary>
    public class HeaderInspector
    {
        public static void InspectFile(string excelPath)
        {
            var excelManager = new ExcelManager();
            var workbook = excelManager.OpenWorkbook(excelPath);
            
            var sheetNames = excelManager.GetSheetNames(workbook);
            Console.WriteLine($"Found {sheetNames.Count} sheets:");
            foreach (var name in sheetNames)
            {
                Console.WriteLine($"  - '{name}'");
            }
            
            int nextSheetNumber = excelManager.GetNextSheetNumber(sheetNames);
            Console.WriteLine($"\nNext sheet number would be: {nextSheetNumber}");
            
            int lastSheetNumber = nextSheetNumber - 1;
            Console.WriteLine($"Looking for sheet: '{lastSheetNumber}'");
            
            var lastSheet = excelManager.GetSheetByName(workbook, lastSheetNumber.ToString());
            if (lastSheet != null)
            {
                string header = excelManager.ReadHeader(lastSheet);
                Console.WriteLine($"\nHeader from sheet '{lastSheetNumber}':");
                Console.WriteLine($"'{header}'");
                Console.WriteLine($"\nLength: {header.Length} characters");
                
                // Show character by character for debugging
                Console.WriteLine("\nCharacter breakdown:");
                for (int i = 0; i < Math.Min(header.Length, 100); i++)
                {
                    Console.WriteLine($"  [{i}] = '{header[i]}' (U+{((int)header[i]):X4})");
                }
            }
            else
            {
                Console.WriteLine($"\nSheet '{lastSheetNumber}' not found!");
            }
            
            // Also check sheet '5' and output.xlsx
            Console.WriteLine("\n\n=== Checking sheet '5' in detail ===");
            var sheet5 = excelManager.GetSheetByName(workbook, "5");
            if (sheet5 != null)
            {
                string header5 = excelManager.ReadHeader(sheet5);
                Console.WriteLine($"\nHeader from sheet '5' (A1):");
                Console.WriteLine($"'{header5}'");
                
                // Check if it's a merged cell
                var ws = sheet5.Worksheet;
                var cell = ws.Cells[1, 1];
                Console.WriteLine($"\nCell A1 info:");
                Console.WriteLine($"  Merge: {cell.Merge}");
                Console.WriteLine($"  Value: {cell.Value}");
                Console.WriteLine($"  Text: {cell.Text}");
                Console.WriteLine($"  Formula: {cell.Formula}");
                Console.WriteLine($"  Font: {cell.Style.Font.Name}, Size: {cell.Style.Font.Size}, Bold: {cell.Style.Font.Bold}");
                
                // Check first few cells in row 1
                Console.WriteLine($"\nFirst 5 cells in row 1:");
                for (int col = 1; col <= 5; col++)
                {
                    var c = ws.Cells[1, col];
                    Console.WriteLine($"  [{col}] Value='{c.Value}', Text='{c.Text}', Formula='{c.Formula}'");
                }
                
                // Check row 2 column headers
                Console.WriteLine($"\nRow 2 (column headers):");
                for (int col = 1; col <= 15; col++)
                {
                    var c = ws.Cells[2, col];
                    if (!string.IsNullOrEmpty(c.Text))
                    {
                        Console.WriteLine($"  [{col}] '{c.Text}'");
                    }
                }
            }
            else
            {
                Console.WriteLine("Sheet '5' not found!");
            }
        }
    }
}
