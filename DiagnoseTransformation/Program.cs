using System;
using System.IO;
using OfficeOpenXml;

namespace DiagnoseTransformation;

class Program
{
    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        
        string excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
        
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"File not found: {excelPath}");
            return;
        }
        
        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            var fissiSheet = package.Workbook.Worksheets["fissi"];
            if (fissiSheet == null)
            {
                Console.WriteLine("Sheet 'fissi' not found");
                return;
            }
            
            Console.WriteLine("=== ANALYZING FISSI SHEET TIME VALUES ===\n");
            
            // Check row 3, columns 2 and 9
            for (int row = 3; row <= 5; row++)
            {
                Console.WriteLine($"Row {row}:");
                
                // Column 2 (Partenza)
                var cell2 = fissiSheet.Cells[row, 2];
                Console.WriteLine($"  Column 2 (Partenza):");
                Console.WriteLine($"    Value: {cell2.Value}");
                Console.WriteLine($"    Value Type: {cell2.Value?.GetType().FullName ?? "null"}");
                Console.WriteLine($"    Text: {cell2.Text}");
                Console.WriteLine($"    Format: {cell2.Style.Numberformat.Format}");
                
                // Test conversion logic
                var sourceValue = cell2.Value;
                if (sourceValue is double || sourceValue is decimal)
                {
                    Console.WriteLine($"    → Detected as double/decimal: {sourceValue}");
                }
                else if (sourceValue is DateTime dt)
                {
                    Console.WriteLine($"    → Detected as DateTime: {dt}");
                    Console.WriteLine($"    → TimeOfDay: {dt.TimeOfDay}");
                    Console.WriteLine($"    → TotalDays: {dt.TimeOfDay.TotalDays}");
                }
                else if (sourceValue is string strValue)
                {
                    Console.WriteLine($"    → Detected as string: '{strValue}'");
                }
                else
                {
                    Console.WriteLine($"    → Other type");
                }
                
                Console.WriteLine();
                
                // Column 9 (Arrivo)
                var cell9 = fissiSheet.Cells[row, 9];
                Console.WriteLine($"  Column 9 (Arrivo):");
                Console.WriteLine($"    Value: {cell9.Value}");
                Console.WriteLine($"    Value Type: {cell9.Value?.GetType().FullName ?? "null"}");
                Console.WriteLine($"    Text: {cell9.Text}");
                Console.WriteLine($"    Format: {cell9.Style.Numberformat.Format}");
                
                sourceValue = cell9.Value;
                if (sourceValue is double || sourceValue is decimal)
                {
                    Console.WriteLine($"    → Detected as double/decimal: {sourceValue}");
                }
                else if (sourceValue is DateTime dt)
                {
                    Console.WriteLine($"    → Detected as DateTime: {dt}");
                    Console.WriteLine($"    → TimeOfDay: {dt.TimeOfDay}");
                    Console.WriteLine($"    → TotalDays: {dt.TimeOfDay.TotalDays}");
                }
                else if (sourceValue is string strValue)
                {
                    Console.WriteLine($"    → Detected as string: '{strValue}'");
                }
                else
                {
                    Console.WriteLine($"    → Other type");
                }
                
                Console.WriteLine("\n" + new string('-', 60) + "\n");
            }
        }
        
        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }
}
