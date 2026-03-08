using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Diagnostic test to inspect the real Excel file and understand the bug.
    /// </summary>
    [TestFixture]
    public class DiagnoseRealFileBug
    {
        [Test]
        public void InspectRealExcelFile()
        {
            string excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
            
            if (!File.Exists(excelPath))
            {
                Assert.Fail($"Excel file not found: {excelPath}");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                Console.WriteLine("\n=== INSPECTING REAL EXCEL FILE ===");
                Console.WriteLine($"File: {excelPath}");
                Console.WriteLine($"\nWorksheets in file:");
                
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - '{ws.Name}' (Dimension: {ws.Dimension?.Address ?? "NULL"})");
                }

                // Try to find laboratori sheet
                var laboratoriSheet = package.Workbook.Worksheets["laboratori"];
                
                if (laboratoriSheet == null)
                {
                    Console.WriteLine("\n❌ NO 'laboratori' SHEET FOUND IN FILE");
                    Console.WriteLine("This explains why no laboratori data appears in output!");
                }
                else
                {
                    Console.WriteLine("\n✓ 'laboratori' sheet found");
                    Console.WriteLine($"  Dimension: {laboratoriSheet.Dimension?.Address ?? "NULL"}");
                    
                    if (laboratoriSheet.Dimension != null)
                    {
                        Console.WriteLine($"\n  First 5 rows:");
                        for (int row = 1; row <= Math.Min(5, laboratoriSheet.Dimension.End.Row); row++)
                        {
                            Console.Write($"    Row {row}: ");
                            for (int col = 1; col <= Math.Min(10, laboratoriSheet.Dimension.End.Column); col++)
                            {
                                var value = laboratoriSheet.Cells[row, col].Text;
                                Console.Write($"[{value}] ");
                            }
                            Console.WriteLine();
                        }
                    }
                }
            }
        }

        [Test]
        public void InspectOutputFile()
        {
            string outputPath = "output_test_bug.xlsx";
            
            if (!File.Exists(outputPath))
            {
                Assert.Inconclusive($"Output file not found: {outputPath}. Run the application first.");
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(outputPath)))
            {
                Console.WriteLine("\n=== INSPECTING OUTPUT FILE ===");
                Console.WriteLine($"File: {outputPath}");
                Console.WriteLine($"\nWorksheets in file:");
                
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - '{ws.Name}' (Dimension: {ws.Dimension?.Address ?? "NULL"})");
                }

                // Find the numbered sheet (should be "2")
                var outputSheet = package.Workbook.Worksheets["2"];
                
                if (outputSheet == null)
                {
                    Console.WriteLine("\n❌ NO numbered sheet found in output");
                    return;
                }

                Console.WriteLine($"\n✓ Sheet '2' found");
                Console.WriteLine($"  Dimension: {outputSheet.Dimension?.Address ?? "NULL"}");
                
                if (outputSheet.Dimension != null)
                {
                    int totalRows = outputSheet.Dimension.End.Row;
                    Console.WriteLine($"  Total rows: {totalRows}");
                    
                    // Count laboratori records (rows with Avv column data)
                    int laboratoriCount = 0;
                    for (int row = 3; row <= totalRows; row++)
                    {
                        var avvValue = outputSheet.Cells[row, 10].Text?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(avvValue) && avvValue.StartsWith("Avv"))
                        {
                            laboratoriCount++;
                        }
                    }
                    
                    Console.WriteLine($"  Laboratori records found: {laboratoriCount}");
                    
                    if (laboratoriCount == 0)
                    {
                        Console.WriteLine("\n❌ BUG CONFIRMED: No laboratori records in output");
                    }
                    else
                    {
                        Console.WriteLine($"\n✓ Found {laboratoriCount} laboratori records");
                    }
                }
            }
        }
    }
}
