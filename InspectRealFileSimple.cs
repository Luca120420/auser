using System;
using System.IO;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        string excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
        
        if (!File.Exists(excelPath))
        {
            Console.WriteLine($"File not found: {excelPath}");
            return;
        }

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            Console.WriteLine("=== REAL FILE INSPECTION ===");
            Console.WriteLine($"File: {excelPath}\n");
            Console.WriteLine("All worksheets:");
            
            foreach (var ws in package.Workbook.Worksheets)
            {
                Console.WriteLine($"  - '{ws.Name}'");
            }

            Console.WriteLine("\nSearching for 'laboratori' sheet...");
            var laboratoriSheet = package.Workbook.Worksheets["laboratori"];
            
            if (laboratoriSheet == null)
            {
                Console.WriteLine("❌ NO 'laboratori' sheet found!");
            }
            else
            {
                Console.WriteLine("✓ 'laboratori' sheet EXISTS");
                
                if (laboratoriSheet.Dimension != null)
                {
                    Console.WriteLine($"  Dimension: {laboratoriSheet.Dimension.Address}");
                    Console.WriteLine($"  Row 2, Col 1: '{laboratoriSheet.Cells[2, 1].Text}'");
                    
                    int dataRows = 0;
                    for (int row = 3; row <= Math.Min(15, laboratoriSheet.Dimension.End.Row); row++)
                    {
                        var dataVal = laboratoriSheet.Cells[row, 1].Text;
                        if (!string.IsNullOrWhiteSpace(dataVal))
                        {
                            dataRows++;
                            Console.WriteLine($"    Row {row}: Data='{dataVal}'");
                        }
                    }
                    Console.WriteLine($"  Total data rows with non-empty Data: {dataRows}");
                }
            }
        }
    }
}
