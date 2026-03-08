using OfficeOpenXml;
using System;
using System.IO;

ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

string filePath = @"..\2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";

if (!File.Exists(filePath))
{
    Console.WriteLine($"ERROR: File not found: {filePath}");
    return;
}

Console.WriteLine($"Inspecting file: {Path.GetFileName(filePath)}");
Console.WriteLine();

using (var package = new ExcelPackage(new FileInfo(filePath)))
{
    Console.WriteLine($"Total worksheets: {package.Workbook.Worksheets.Count}");
    Console.WriteLine();
    Console.WriteLine("Sheet names:");
    Console.WriteLine("============");
    
    bool hasLaboratori = false;
    
    foreach (var worksheet in package.Workbook.Worksheets)
    {
        Console.WriteLine($"  - '{worksheet.Name}' (length: {worksheet.Name.Length})");
        
        if (worksheet.Name.Equals("laboratori", StringComparison.OrdinalIgnoreCase) ||
            worksheet.Name.Trim().Equals("laboratori", StringComparison.OrdinalIgnoreCase))
        {
            hasLaboratori = true;
            Console.WriteLine($"    ^ FOUND LABORATORI SHEET (with {(worksheet.Name != worksheet.Name.Trim() ? "TRAILING SPACE" : "exact match")})!");
            Console.WriteLine($"    Dimension: {(worksheet.Dimension != null ? worksheet.Dimension.Address : "NULL (empty sheet)")}");
            
            if (worksheet.Dimension != null)
            {
                Console.WriteLine($"    Rows: {worksheet.Dimension.End.Row}");
                Console.WriteLine($"    Columns: {worksheet.Dimension.End.Column}");
                
                // Check row 1 and row 2
                var a1 = worksheet.Cells[1, 1].Text;
                var a2 = worksheet.Cells[2, 1].Text;
                
                Console.WriteLine($"    Cell A1: '{a1}'");
                Console.WriteLine($"    Cell A2: '{a2}'");
            }
        }
    }
    
    Console.WriteLine();
    Console.WriteLine($"Result: {(hasLaboratori ? "YES - 'laboratori' sheet EXISTS" : "NO - 'laboratori' sheet NOT FOUND")}");
}
