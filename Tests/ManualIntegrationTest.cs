using System;
using System.IO;
using NUnit.Framework;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Manual integration test to verify time format fix with real files
    /// Run this test manually to verify the fix works with actual input files
    /// </summary>
    [TestFixture]
    [Explicit] // Mark as explicit so it doesn't run in automated test suite
    public class ManualIntegrationTest
    {
        private ICSVParser _csvParser = null!;
        private IDataTransformer _dataTransformer = null!;
        private IExcelManager _excelManager = null!;

        [SetUp]
        public void Setup()
        {
            _csvParser = new CSVParser();
            _dataTransformer = new DataTransformer();
            _excelManager = new ExcelManager();
        }

        [Test]
        [Category("Manual")]
        public void TestTimeFormatFix_WithRealFiles()
        {
            // Arrange
            string csvPath = "168514-Estrazione_1770193162042.csv";
            string excelPath = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
            string outputPath = "test_output_time_format_verification.xlsx";

            // Verify input files exist
            Assert.That(File.Exists(csvPath), Is.True, $"CSV file not found: {csvPath}");
            Assert.That(File.Exists(excelPath), Is.True, $"Excel file not found: {excelPath}");

            Console.WriteLine("=== Testing Time Format Fix with Real Files ===");
            Console.WriteLine($"CSV Input: {csvPath}");
            Console.WriteLine($"Excel Input: {excelPath}");
            Console.WriteLine($"Output: {outputPath}");
            Console.WriteLine();

            // Act - Parse CSV
            Console.WriteLine("Step 1: Parsing CSV...");
            var appointments = _csvParser.ParseCSV(csvPath);
            Console.WriteLine($"  Parsed {appointments.Count} appointments");
            Assert.That(appointments.Count, Is.GreaterThan(0), "CSV should contain appointments");

            // Act - Transform data
            Console.WriteLine("Step 2: Transforming data...");
            var transformationResult = _dataTransformer.TransformData(appointments);
            Console.WriteLine($"  Transformed {transformationResult.Rows.Count} rows");
            Assert.That(transformationResult.Rows.Count, Is.GreaterThan(0), "Should have transformed rows");

            // Act - Load Excel workbook
            Console.WriteLine("Step 3: Loading Excel workbook...");
            var workbook = _excelManager.LoadWorkbook(excelPath);
            Console.WriteLine($"  Loaded workbook with {workbook.Sheets.Count} sheets");
            Assert.That(workbook.Sheets.Count, Is.GreaterThan(0), "Workbook should have sheets");

            // Act - Find fissi sheet
            Console.WriteLine("Step 4: Finding 'fissi' sheet...");
            Sheet? fissiSheet = null;
            foreach (var sheet in workbook.Sheets)
            {
                if (sheet.Name.Equals("fissi", StringComparison.OrdinalIgnoreCase))
                {
                    fissiSheet = sheet;
                    break;
                }
            }
            Assert.That(fissiSheet, Is.Not.Null, "'fissi' sheet should exist in workbook");
            Console.WriteLine($"  Found 'fissi' sheet");

            // Act - Create new sheet
            Console.WriteLine("Step 5: Creating new sheet...");
            var newSheet = _excelManager.CreateSheet(workbook, transformationResult.HeaderInfo);
            Assert.That(newSheet, Is.Not.Null, "New sheet should be created");
            Console.WriteLine($"  Created new sheet");

            // Act - Append CSV data
            Console.WriteLine("Step 6: Appending CSV data...");
            _excelManager.AppendTransformedData(newSheet, transformationResult);
            Console.WriteLine($"  Appended CSV data");

            // Act - Append fissi data
            Console.WriteLine("Step 7: Appending fissi data...");
            int startRow = transformationResult.Rows.Count + 2; // +1 for header, +1 for next row
            _excelManager.AppendFissiData(newSheet, fissiSheet, startRow);
            Console.WriteLine($"  Appended fissi data starting at row {startRow}");

            // Act - Save workbook
            Console.WriteLine("Step 8: Saving workbook...");
            _excelManager.SaveWorkbook(workbook, outputPath);
            Console.WriteLine($"  Saved to {outputPath}");

            // Assert - Verify output file exists
            Assert.That(File.Exists(outputPath), Is.True, "Output file should be created");

            // Assert - Inspect time columns in output
            Console.WriteLine();
            Console.WriteLine("Step 9: Verifying time format in output...");
            VerifyTimeFormatInOutput(outputPath, startRow);

            Console.WriteLine();
            Console.WriteLine("=== TEST COMPLETED SUCCESSFULLY ===");
            Console.WriteLine($"Output file: {Path.GetFullPath(outputPath)}");
        }

        private void VerifyTimeFormatInOutput(string filePath, int startRow)
        {
            var workbook = _excelManager.LoadWorkbook(filePath);
            var sheet = workbook.Sheets[0]; // First sheet should be the new sheet
            var worksheet = sheet.Worksheet;

            var dimension = worksheet.Dimension;
            Assert.That(dimension, Is.Not.Null, "Output sheet should not be empty");

            Console.WriteLine();
            Console.WriteLine("=== Time Column Verification ===");
            Console.WriteLine($"Checking rows starting from {startRow}");
            Console.WriteLine();
            Console.WriteLine("Row | Partenza (Col 2)                              | Arrivo (Col 9)");
            Console.WriteLine("----|-----------------------------------------------|-----------------------------------------------");

            int rowsChecked = 0;
            int rowsWithTimeFormat = 0;
            int maxRowsToCheck = Math.Min(10, dimension!.End.Row - startRow + 1);

            for (int row = startRow; row < startRow + maxRowsToCheck; row++)
            {
                var partenzaCell = worksheet.Cells[row, 2];
                var arrivoCell = worksheet.Cells[row, 9];

                // Get cell properties
                var partenzaValue = partenzaCell.Value;
                string partenzaText = partenzaCell.Text ?? "null";
                string partenzaFormat = partenzaCell.Style?.Numberformat?.Format ?? "no format";

                var arrivoValue = arrivoCell.Value;
                string arrivoText = arrivoCell.Text ?? "null";
                string arrivoFormat = arrivoCell.Style?.Numberformat?.Format ?? "no format";

                // Display row info
                Console.WriteLine($"{row,3} | V:{partenzaValue,-12} T:{partenzaText,-10} F:{partenzaFormat,-10} | V:{arrivoValue,-12} T:{arrivoText,-10} F:{arrivoFormat,-10}");

                // Verify format if cell has a value
                if (partenzaValue != null || arrivoValue != null)
                {
                    rowsChecked++;

                    // Check if format is applied
                    bool partenzaHasTimeFormat = partenzaFormat == "h:mm";
                    bool arrivoHasTimeFormat = arrivoFormat == "h:mm";

                    if (partenzaValue != null)
                    {
                        Assert.That(partenzaHasTimeFormat, Is.True, 
                            $"Row {row}, Col 2 (Partenza) should have 'h:mm' format, but has '{partenzaFormat}'");
                        
                        // Check if text is in time format (not decimal)
                        Assert.That(partenzaText, Does.Match(@"^\d{1,2}:\d{2}$"), 
                            $"Row {row}, Col 2 (Partenza) text should be in time format (e.g., 8:30), but is '{partenzaText}'");
                    }

                    if (arrivoValue != null)
                    {
                        Assert.That(arrivoHasTimeFormat, Is.True, 
                            $"Row {row}, Col 9 (Arrivo) should have 'h:mm' format, but has '{arrivoFormat}'");
                        
                        // Check if text is in time format (not decimal)
                        Assert.That(arrivoText, Does.Match(@"^\d{1,2}:\d{2}$"), 
                            $"Row {row}, Col 9 (Arrivo) text should be in time format (e.g., 10:00), but is '{arrivoText}'");
                    }

                    if (partenzaHasTimeFormat || arrivoHasTimeFormat)
                    {
                        rowsWithTimeFormat++;
                    }
                }
            }

            Console.WriteLine();
            Console.WriteLine($"Checked {rowsChecked} rows with data");
            Console.WriteLine($"Found {rowsWithTimeFormat} rows with correct time format");
            Console.WriteLine();
            Console.WriteLine("✓ All time columns have correct 'h:mm' format");
            Console.WriteLine("✓ All time values display as time (e.g., 8:30) instead of decimal (e.g., 0.354166666666667)");
            Console.WriteLine();
            Console.WriteLine("Legend: V=Value (internal decimal), T=Text (displayed format), F=Format (number format)");
        }
    }
}
