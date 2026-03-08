using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer
{
    class TestLaboratoriProcessing
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Find the input Excel file
            var inputFile = Directory.GetFiles(".", "*.xlsx")
                .Where(f => !f.Contains("~$") && !f.Contains("output"))
                .OrderByDescending(f => File.GetLastWriteTime(f))
                .FirstOrDefault();
            
            if (inputFile == null)
            {
                Console.WriteLine("No input Excel file found");
                return;
            }
            
            Console.WriteLine($"Testing with file: {Path.GetFileName(inputFile)}");
            Console.WriteLine();
            
            using (var package = new ExcelPackage(new FileInfo(inputFile)))
            {
                var workbook = new Models.ExcelWorkbook(package);
                var excelManager = new ExcelManager();
                
                // Check for laboratori sheet
                Console.WriteLine("=== Checking for laboratori sheet ===");
                var laboratoriSheet = excelManager.GetSheetByName(workbook, "laboratori");
                
                if (laboratoriSheet == null)
                {
                    Console.WriteLine("❌ laboratori sheet NOT FOUND");
                    Console.WriteLine($"Available sheets: {string.Join(", ", package.Workbook.Worksheets.Select(ws => ws.Name))}");
                    return;
                }
                
                Console.WriteLine("✓ laboratori sheet FOUND");
                Console.WriteLine();
                
                // Check sheet dimensions
                var dimension = laboratoriSheet.Worksheet.Dimension;
                if (dimension == null)
                {
                    Console.WriteLine("❌ laboratori sheet is EMPTY (no dimension)");
                    return;
                }
                
                Console.WriteLine($"Sheet dimensions: {dimension.Address}");
                Console.WriteLine($"Rows: {dimension.Start.Row} to {dimension.End.Row}");
                Console.WriteLine($"Columns: {dimension.Start.Column} to {dimension.End.Column}");
                Console.WriteLine();
                
                // Check row 1 and 2
                Console.WriteLine("=== Header Detection ===");
                var row1Col1 = laboratoriSheet.Worksheet.Cells[1, 1].Text?.Trim() ?? "";
                var row2Col1 = laboratoriSheet.Worksheet.Cells[2, 1].Text?.Trim() ?? "";
                
                Console.WriteLine($"Row 1, Col 1: '{row1Col1}'");
                Console.WriteLine($"Row 2, Col 1: '{row2Col1}'");
                Console.WriteLine($"Row 2 contains 'Data': {row2Col1.Equals("Data", StringComparison.OrdinalIgnoreCase)}");
                Console.WriteLine();
                
                // Determine data start row
                int dataStartRow = 3;
                if (row2Col1.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    dataStartRow = 3;
                    Console.WriteLine("✓ Headers in row 2, data starts at row 3");
                }
                else if (row1Col1.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    dataStartRow = 2;
                    Console.WriteLine("✓ Headers in row 1, data starts at row 2");
                }
                else
                {
                    dataStartRow = 2;
                    Console.WriteLine("⚠ No 'Data' header found, assuming data starts at row 2");
                }
                Console.WriteLine();
                
                // Count data rows
                Console.WriteLine("=== Data Rows ===");
                int validDataRows = 0;
                int emptyDataRows = 0;
                
                for (int row = dataStartRow; row <= dimension.End.Row; row++)
                {
                    var dataValue = laboratoriSheet.Worksheet.Cells[row, 1].Value;
                    var dataText = laboratoriSheet.Worksheet.Cells[row, 1].Text?.Trim() ?? "";
                    
                    if (dataValue == null || string.IsNullOrWhiteSpace(dataText))
                    {
                        emptyDataRows++;
                        Console.WriteLine($"Row {row}: EMPTY (will be skipped)");
                    }
                    else
                    {
                        validDataRows++;
                        var partenza = laboratoriSheet.Worksheet.Cells[row, 2].Text;
                        var assistito = laboratoriSheet.Worksheet.Cells[row, 3].Text;
                        var avv = laboratoriSheet.Worksheet.Cells[row, 10].Text;
                        Console.WriteLine($"Row {row}: Data='{dataText}', Partenza='{partenza}', Assistito='{assistito}', Avv='{avv}'");
                    }
                }
                
                Console.WriteLine();
                Console.WriteLine($"Valid data rows: {validDataRows}");
                Console.WriteLine($"Empty data rows (skipped): {emptyDataRows}");
                Console.WriteLine($"Expected rows in output: {validDataRows}");
                Console.WriteLine();
                
                // Test AppendLaboratoriData
                Console.WriteLine("=== Testing AppendLaboratoriData ===");
                
                // Create a test output sheet
                var testOutputSheet = package.Workbook.Worksheets.Add("TestOutput");
                var targetSheet = new Sheet(testOutputSheet);
                
                try
                {
                    excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1);
                    
                    var outputDimension = testOutputSheet.Dimension;
                    if (outputDimension == null)
                    {
                        Console.WriteLine("❌ NO DATA was written to output sheet");
                    }
                    else
                    {
                        int rowsWritten = outputDimension.End.Row;
                        Console.WriteLine($"✓ {rowsWritten} rows written to output sheet");
                        
                        // Check for Avv column data
                        int avvCount = 0;
                        for (int row = 1; row <= rowsWritten; row++)
                        {
                            var avvValue = testOutputSheet.Cells[row, 10].Text;
                            if (!string.IsNullOrWhiteSpace(avvValue))
                            {
                                avvCount++;
                                Console.WriteLine($"  Row {row}: Avv='{avvValue}'");
                            }
                        }
                        Console.WriteLine($"Rows with Avv data: {avvCount}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"❌ ERROR: {ex.Message}");
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                }
            }
        }
    }
}
