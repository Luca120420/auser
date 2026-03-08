using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    [TestFixture]
    public class InspectRealFile
    {
        [Test]
        public void CheckForLaboratoriSheet()
        {
            string excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
            
            if (!File.Exists(excelPath))
            {
                Assert.Inconclusive($"File not found: {excelPath}");
                return;
            }

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                Console.WriteLine("\n=== INSPECTING REAL FILE ===");
                Console.WriteLine($"File: {excelPath}\n");
                Console.WriteLine("All worksheets in file:");
                
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - '{ws.Name}' (Dimension: {ws.Dimension?.Address ?? "NULL"})");
                }

                Console.WriteLine("\nSearching for 'laboratori' sheet...");
                var laboratoriSheet = package.Workbook.Worksheets["laboratori"];
                
                if (laboratoriSheet == null)
                {
                    Console.WriteLine("❌ NO 'laboratori' sheet found in file!");
                    Console.WriteLine("\nThis explains why no laboratori data appears - the sheet doesn't exist!");
                    Assert.Pass("Confirmed: laboratori sheet does NOT exist in real file");
                }
                else
                {
                    Console.WriteLine("✓ 'laboratori' sheet FOUND");
                    
                    if (laboratoriSheet.Dimension != null)
                    {
                        Console.WriteLine($"  Dimension: {laboratoriSheet.Dimension.Address}");
                        Console.WriteLine($"  Row 2, Col 1: '{laboratoriSheet.Cells[2, 1].Text}'");
                        
                        int dataRows = 0;
                        for (int row = 3; row <= laboratoriSheet.Dimension.End.Row; row++)
                        {
                            if (!string.IsNullOrWhiteSpace(laboratoriSheet.Cells[row, 1].Text))
                            {
                                dataRows++;
                            }
                        }
                        Console.WriteLine($"  Data rows with non-empty Data: {dataRows}");
                        
                        Assert.Pass($"Laboratori sheet exists with {dataRows} data rows");
                    }
                    else
                    {
                        Console.WriteLine("  Sheet is empty (Dimension is NULL)");
                        Assert.Pass("Laboratori sheet exists but is empty");
                    }
                }
            }
        }
    }
}
