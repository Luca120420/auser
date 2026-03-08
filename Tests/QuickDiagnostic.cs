using System;
using System.IO;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    [TestFixture]
    public class QuickDiagnostic
    {
        [Test]
        public void CheckRealFile()
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
                Console.WriteLine("\n=== REAL FILE INSPECTION ===");
                Console.WriteLine($"File: {excelPath}");
                Console.WriteLine($"\nWorksheets:");
                
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - '{ws.Name}'");
                }

                var laboratoriSheet = package.Workbook.Worksheets["laboratori"];
                
                if (laboratoriSheet == null)
                {
                    Console.WriteLine("\n❌ NO 'laboratori' sheet found!");
                    Assert.Fail("Laboratori sheet not found in real file");
                }
                else
                {
                    Console.WriteLine("\n✓ 'laboratori' sheet EXISTS");
                    
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
                    }
                }
            }
        }
    }
}
