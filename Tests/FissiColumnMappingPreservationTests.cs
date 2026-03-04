using System;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Preservation property tests for fissi column mapping fix.
    /// 
    /// IMPORTANT: These tests verify behaviors that should remain UNCHANGED.
    /// These tests should PASS on unfixed code and continue to PASS after the fix.
    /// 
    /// This follows the observation-first methodology:
    /// 1. Observe behavior on UNFIXED code for non-buggy inputs
    /// 2. Write property-based tests capturing observed behavior patterns
    /// 3. Run tests on UNFIXED code - they should PASS
    /// 4. After fix is implemented, re-run tests - they should still PASS
    /// 
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, 3.9, 3.10**
    /// </summary>
    [TestFixture]
    public class FissiColumnMappingPreservationTests
    {
        private ExcelManager _excelManager = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Property 4: Preservation - Columns 1-3 Mapping
        /// 
        /// For any fissi sheet data where columns 1-3 contain values (Data, Partenza, Assistito),
        /// the AppendFissiData method SHALL continue to map these columns to target columns 1-3
        /// respectively, preserving the existing correct behavior.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1, 3.2, 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_Columns1to3Mapping()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                ArbitraryFissiColumns1to3Data(),
                (FissiColumns1to3Data data) =>
                {
                    // Arrange - Create fissi sheet with data in columns 1-3
                    using (var package = new ExcelPackage())
                    {
                        var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                        var targetWorksheet = package.Workbook.Worksheets.Add("target");

                        // Add header row (row 2 contains "Data" as per header detection logic)
                        fissiWorksheet.Cells[2, 1].Value = "Data";
                        fissiWorksheet.Cells[2, 2].Value = "Partenza";
                        fissiWorksheet.Cells[2, 3].Value = "Assistito";

                        // Add data row with values in columns 1-3
                        fissiWorksheet.Cells[3, 1].Value = data.Data;
                        fissiWorksheet.Cells[3, 2].Value = data.Partenza;
                        fissiWorksheet.Cells[3, 3].Value = data.Assistito;

                        var fissiSheet = new Sheet(fissiWorksheet);
                        var targetSheet = new Sheet(targetWorksheet);

                        // Act
                        _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                        // Assert - Verify columns 1-3 map correctly
                        var dataValue = targetWorksheet.Cells[1, 1].Value?.ToString();
                        var partenzaValue = targetWorksheet.Cells[1, 2].Value?.ToString();
                        var assistitoValue = targetWorksheet.Cells[1, 3].Value?.ToString();

                        return dataValue == data.Data &&
                               partenzaValue == data.Partenza &&
                               assistitoValue == data.Assistito;
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 5: Preservation - Header Detection Logic
        /// 
        /// For any fissi sheet with header rows, the AppendFissiData method SHALL continue to
        /// detect and skip header rows correctly (checking for "Data" in rows 1-2).
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_HeaderDetection()
        {
            // Test case 1: Header in row 2 (most common case)
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row in row 2
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";
                fissiWorksheet.Cells[2, 3].Value = "Assistito";

                // Add data row in row 3
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 2].Value = "09:00";
                fissiWorksheet.Cells[3, 3].Value = "Rossi Mario";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Header row should be skipped, data should start in target row 1
                Assert.That(targetWorksheet.Cells[1, 1].Value?.ToString(), Is.EqualTo("15/01/2024"),
                    "Header in row 2 should be skipped, data should start in target row 1");
            }

            // Test case 2: Header in row 1
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row in row 1
                fissiWorksheet.Cells[1, 1].Value = "Data";
                fissiWorksheet.Cells[1, 2].Value = "Partenza";
                fissiWorksheet.Cells[1, 3].Value = "Assistito";

                // Add data row in row 2
                fissiWorksheet.Cells[2, 1].Value = "16/01/2024";
                fissiWorksheet.Cells[2, 2].Value = "10:00";
                fissiWorksheet.Cells[2, 3].Value = "Bianchi Luigi";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Header row should be skipped, data should start in target row 1
                Assert.That(targetWorksheet.Cells[1, 1].Value?.ToString(), Is.EqualTo("16/01/2024"),
                    "Header in row 1 should be skipped, data should start in target row 1");
            }

            // Test case 3: No header row (data starts in row 2)
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // No header row, data starts in row 2
                fissiWorksheet.Cells[2, 1].Value = "17/01/2024";
                fissiWorksheet.Cells[2, 2].Value = "11:00";
                fissiWorksheet.Cells[2, 3].Value = "Verdi Giuseppe";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Data should be copied starting from target row 1
                Assert.That(targetWorksheet.Cells[1, 1].Value?.ToString(), Is.EqualTo("17/01/2024"),
                    "When no header row, data should start in target row 1");
            }
        }

        /// <summary>
        /// Property 5: Preservation - Cell Formatting Copy Logic
        /// 
        /// For any fissi sheet data, the AppendFissiData method SHALL continue to copy
        /// cell formatting (fonts, borders, number formats) correctly, while skipping
        /// background colors.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.6, 3.7, 3.8, 3.9**
        /// </summary>
        [Test]
        public void Property_Preservation_CellFormattingCopy()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";

                // Add data row with formatting in column 1
                var sourceCell = fissiWorksheet.Cells[3, 1];
                sourceCell.Value = "15/01/2024";
                
                // Apply various formatting
                sourceCell.Style.Font.Bold = true;
                sourceCell.Style.Font.Italic = true;
                sourceCell.Style.Font.Size = 14;
                sourceCell.Style.Font.Name = "Arial";
                sourceCell.Style.Font.Color.SetColor(System.Drawing.Color.Red);
                
                // Apply background color (should NOT be copied)
                sourceCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                sourceCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                
                // Apply borders
                sourceCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sourceCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sourceCell.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sourceCell.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                
                // Apply number format
                sourceCell.Style.Numberformat.Format = "dd/mm/yyyy";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Verify formatting is copied correctly
                var targetCell = targetWorksheet.Cells[1, 1];
                
                // Font properties should be copied
                Assert.That(targetCell.Style.Font.Bold, Is.True, "Font bold should be copied");
                Assert.That(targetCell.Style.Font.Italic, Is.True, "Font italic should be copied");
                Assert.That(targetCell.Style.Font.Size, Is.EqualTo(14), "Font size should be copied");
                Assert.That(targetCell.Style.Font.Name, Is.EqualTo("Arial"), "Font name should be copied");
                
                // Background color should NOT be copied (requirement 3.7)
                Assert.That(targetCell.Style.Fill.PatternType, Is.Not.EqualTo(ExcelFillStyle.Solid),
                    "Background color should NOT be copied to avoid yellow highlighting");
                
                // Borders should be copied
                Assert.That(targetCell.Style.Border.Top.Style, Is.EqualTo(ExcelBorderStyle.Thin),
                    "Top border should be copied");
                Assert.That(targetCell.Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thin),
                    "Bottom border should be copied");
                
                // Number format should be copied
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("dd/mm/yyyy"),
                    "Number format should be copied");
            }
        }

        /// <summary>
        /// Property 5: Preservation - Null Validation
        /// 
        /// For any null targetSheet or fissiSheet, the AppendFissiData method SHALL continue
        /// to throw ArgumentNullException.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.10**
        /// </summary>
        [Test]
        public void Property_Preservation_NullValidation()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("test");
                var validSheet = new Sheet(worksheet);

                // Test null targetSheet
                Assert.Throws<ArgumentNullException>(() =>
                    _excelManager.AppendFissiData(null!, validSheet, 1),
                    "Null targetSheet should throw ArgumentNullException");

                // Test null fissiSheet
                Assert.Throws<ArgumentNullException>(() =>
                    _excelManager.AppendFissiData(validSheet, null!, 1),
                    "Null fissiSheet should throw ArgumentNullException");

                // Test both null
                Assert.Throws<ArgumentNullException>(() =>
                    _excelManager.AppendFissiData(null!, null!, 1),
                    "Both null should throw ArgumentNullException");
            }
        }

        /// <summary>
        /// Property 5: Preservation - Empty Sheet Handling
        /// 
        /// For any empty fissi sheet, the AppendFissiData method SHALL continue to return
        /// without processing (no data copied, no errors thrown).
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.4**
        /// </summary>
        [Test]
        public void Property_Preservation_EmptySheetHandling()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Empty fissi sheet (no data, no dimensions)
                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act - Should return without processing
                Assert.DoesNotThrow(() =>
                    _excelManager.AppendFissiData(targetSheet, fissiSheet, 1),
                    "Empty sheet should not throw exception");

                // Assert - Target sheet should remain empty
                Assert.That(targetWorksheet.Dimension, Is.Null,
                    "Target sheet should remain empty when fissi sheet is empty");
            }
        }

        /// <summary>
        /// Property-based test: Preservation - Multiple Rows with Columns 1-3
        /// 
        /// For any number of rows with data in columns 1-3, the mapping should remain correct.
        /// 
        /// **Validates: Requirements 3.1, 3.2, 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_MultipleRowsColumns1to3()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            Prop.ForAll(
                ArbitraryMultipleFissiRows(),
                (MultipleFissiRowsData data) =>
                {
                    // Arrange
                    using (var package = new ExcelPackage())
                    {
                        var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                        var targetWorksheet = package.Workbook.Worksheets.Add("target");

                        // Add header row
                        fissiWorksheet.Cells[2, 1].Value = "Data";
                        fissiWorksheet.Cells[2, 2].Value = "Partenza";
                        fissiWorksheet.Cells[2, 3].Value = "Assistito";

                        // Add data rows
                        var rowsList = Microsoft.FSharp.Collections.ListModule.ToArray(data.Rows);
                        for (int i = 0; i < rowsList.Length; i++)
                        {
                            var row = rowsList[i];
                            int fissiRow = 3 + i;
                            fissiWorksheet.Cells[fissiRow, 1].Value = row.Data;
                            fissiWorksheet.Cells[fissiRow, 2].Value = row.Partenza;
                            fissiWorksheet.Cells[fissiRow, 3].Value = row.Assistito;
                        }

                        var fissiSheet = new Sheet(fissiWorksheet);
                        var targetSheet = new Sheet(targetWorksheet);

                        // Act
                        _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                        // Assert - Verify all rows are mapped correctly
                        bool allRowsCorrect = true;
                        for (int i = 0; i < rowsList.Length; i++)
                        {
                            var row = rowsList[i];
                            int targetRow = 1 + i;
                            
                            var dataValue = targetWorksheet.Cells[targetRow, 1].Value?.ToString();
                            var partenzaValue = targetWorksheet.Cells[targetRow, 2].Value?.ToString();
                            var assistitoValue = targetWorksheet.Cells[targetRow, 3].Value?.ToString();

                            if (dataValue != row.Data || 
                                partenzaValue != row.Partenza || 
                                assistitoValue != row.Assistito)
                            {
                                allRowsCorrect = false;
                                break;
                            }
                        }

                        return allRowsCorrect;
                    }
                }
            ).Check(config);
        }

        #region Arbitrary Generators

        /// <summary>
        /// Generator for fissi columns 1-3 data (Data, Partenza, Assistito)
        /// </summary>
        private static Arbitrary<FissiColumns1to3Data> ArbitraryFissiColumns1to3Data()
        {
            var gen = from data in Gen.Elements("15/01/2024", "16/01/2024", "17/01/2024", "18/01/2024")
                      from partenza in Gen.Elements("09:00", "10:00", "11:00", "14:00", "15:00")
                      from assistito in Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Giuseppe", "Neri Anna")
                      select new FissiColumns1to3Data
                      {
                          Data = data,
                          Partenza = partenza,
                          Assistito = assistito
                      };

            return Arb.From(gen);
        }

        /// <summary>
        /// Generator for multiple fissi rows (1-5 rows)
        /// </summary>
        private static Arbitrary<MultipleFissiRowsData> ArbitraryMultipleFissiRows()
        {
            var rowGen = from data in Gen.Elements("15/01/2024", "16/01/2024", "17/01/2024")
                         from partenza in Gen.Elements("09:00", "10:00", "11:00")
                         from assistito in Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Giuseppe")
                         select new FissiColumns1to3Data
                         {
                             Data = data,
                             Partenza = partenza,
                             Assistito = assistito
                         };

            var gen = from count in Gen.Choose(1, 5)
                      from rows in Gen.ListOf(count, rowGen)
                      select new MultipleFissiRowsData { Rows = rows };

            return Arb.From(gen);
        }

        #endregion

        #region Test Data Classes

        public class FissiColumns1to3Data
        {
            public string Data { get; set; } = null!;
            public string Partenza { get; set; } = null!;
            public string Assistito { get; set; } = null!;
        }

        public class MultipleFissiRowsData
        {
            public Microsoft.FSharp.Collections.FSharpList<FissiColumns1to3Data> Rows { get; set; } = null!;
        }

        #endregion
    }
}
