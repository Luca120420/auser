using System;
using System.IO;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

class TestHeaderRead
{
    static void Main()
    {
        var excelManager = new ExcelManager();
        var workbook = excelManager.OpenWorkbook("2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx");
        
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
        }
        else
        {
            Console.WriteLine($"\nSheet '{lastSheetNumber}' not found!");
        }
    }
}
