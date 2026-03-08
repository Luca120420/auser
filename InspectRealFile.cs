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
            Console.WriteLine($"❌ File not found: {excelPath}");
            return;
        }

        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (var package = new ExcelPackage(new FileInfo(excelPath)))
        {
            Console.WriteLine($"\n=== INSPECTING: {excelPath} ===\n");
            Console.WriteLine("Worksheets in file:");
            
            bool hasLaboratori = false;
            foreach (var ws in package.Workbook.Worksheets)
            {
                Console.WriteLine($"  - '{ws.Name}'");
                if (ws.Name.Equals("laboratori", StringComparison.OrdinalIgnoreCase))
                {
                    hasLaboratori = true;
                }
            }

            if (!hasLaboratori)
            {
                Console.WriteLine("\n❌ NO 'laboratori' SHEET FOUND");
                Console.WriteLine("This explains why no laboratori data appears in output!");
                Console.WriteLine("The bug is: the real file doesn't have a laboratori sheet.");
            }
            else
            {
                Console.WriteLine("\n✓ 'laboratori' sheet exists");
                var labSheet = package.Workbook.Worksheets["laboratori"];
                Console.WriteLine($"  Dimension: {labSheet.Dimension?.Address ?? "NULL"}");
                
                if (labSheet.Dimension != null)
                {
                    Console.WriteLine($"\n  First 5 rows, first 10 columns:");
                    for (int row = 1; row <= Math.Min(5, labSheet.Dimension.End.Row); row++)
                    {
                        Console.Write($"    Row {row}: ");
                        for (int col = 1; col <= Math.Min(10, labSheet.Dimension.End.Column); col++)
                        {
                            var value = labSheet.Cells[row, col].Text;
                            Console.Write($"[{value}] ");
                        }
                        Console.WriteLine();
                    }
                }
            }
        }
    }
}
