using System;
using System.Collections.Generic;
using System.Linq;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for FormattingService class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 4.1, 10.2
    /// </summary>
    [TestFixture]
    public class FormattingServicePropertyTests
    {
        private FormattingService _formattingService = null!;

        [SetUp]
        public void Setup()
        {
            _formattingService = new FormattingService();
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Helper method to create a test sheet with specified number of columns
        /// </summary>
        private Sheet CreateTestSheet(int columnCount)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");

            // Write header row with specified number of columns
            for (int col = 1; col <= columnCount; col++)
            {
                worksheet.Cells[1, col].Value = $"Column{col}";
            }

            return new Sheet(worksheet);
        }

        // Feature: excel-output-enhancement, Property 3: Header Bold Formatting
        /// <summary>
        /// Property 3: Header Bold Formatting
        /// For any generated Excel output, all cells in the column header row
        /// SHALL have bold formatting applied.
        /// **Validates: Requirements 4.1**
        /// </summary>
        [Test]
        public void Property_HeaderBoldFormatting()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.Choose(1, 50)),
                (int columnCount) =>
                {
                    try
                    {
                        // Arrange - Create a test sheet with random number of columns
                        var sheet = CreateTestSheet(columnCount);
                        var headerRow = 1;

                        // Act - Apply bold formatting to headers
                        _formattingService.ApplyBoldHeaders(sheet, headerRow);

                        // Assert - Verify all header cells have bold formatting
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var cell = sheet.Worksheet.Cells[headerRow, col];
                            if (!cell.Style.Font.Bold)
                            {
                                return false.Label($"Column {col} header is not bold");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Bold formatting failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Helper method to create a test sheet with date groups
        /// </summary>
        private Sheet CreateTestSheetWithDateGroups(List<string> dates)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");

            // Write header row
            worksheet.Cells[1, 1].Value = "Data";
            worksheet.Cells[1, 2].Value = "Column2";

            // Write data rows with dates
            for (int i = 0; i < dates.Count; i++)
            {
                worksheet.Cells[i + 2, 1].Value = dates[i];
                worksheet.Cells[i + 2, 2].Value = $"Value{i}";
            }

            return new Sheet(worksheet);
        }

        /// <summary>
        /// Helper method to identify expected date group boundaries
        /// </summary>
        private List<int> GetExpectedDateGroupBoundaries(List<string> dates, int startRow)
        {
            var boundaries = new List<int>();
            
            for (int i = 0; i < dates.Count; i++)
            {
                // Check if this is the last row or if the next row has a different date
                if (i == dates.Count - 1)
                {
                    boundaries.Add(startRow + i);
                }
                else if (dates[i] != dates[i + 1])
                {
                    boundaries.Add(startRow + i);
                }
            }
            
            return boundaries;
        }

        // Feature: excel-output-enhancement, Property 9: Date Group Border Application
        /// <summary>
        /// Property 9: Date Group Border Application
        /// For any date group in the generated Excel output, the last row of that group
        /// SHALL have a thick bottom border applied.
        /// **Validates: Requirements 10.2**
        /// </summary>
        [Test]
        public void Property_DateGroupBorderApplication()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for date strings
            var dateGen = from year in Gen.Choose(2020, 2030)
                         from month in Gen.Choose(1, 12)
                         from day in Gen.Choose(1, 28)
                         select $"{day:D2}/{month:D2}/{year}";

            // Generator for list of dates (with potential duplicates to create groups)
            var dateListGen = from size in Gen.Choose(1, 20)
                             from dates in Gen.ListOf(size, dateGen)
                             let dateList = dates.ToList()
                             where dateList.Count > 0
                             select dateList;

            Prop.ForAll(
                Arb.From(dateListGen),
                (List<string> dates) =>
                {
                    try
                    {
                        // Arrange - Create a test sheet with date groups
                        var sheet = CreateTestSheetWithDateGroups(dates);
                        var dataStartRow = 2;
                        var dataEndRow = dataStartRow + dates.Count - 1;
                        var dataColumnIndex = 1;

                        // Calculate expected boundaries
                        var expectedBoundaries = GetExpectedDateGroupBoundaries(dates, dataStartRow);

                        // Act - Apply date group borders
                        _formattingService.ApplyDateGroupBorders(sheet, dataStartRow, dataEndRow, dataColumnIndex);

                        // Assert - Verify thick borders are applied to boundary rows
                        for (int row = dataStartRow; row <= dataEndRow; row++)
                        {
                            var cell = sheet.Worksheet.Cells[row, 1];
                            var hasBorder = cell.Style.Border.Bottom.Style == OfficeOpenXml.Style.ExcelBorderStyle.Thick;
                            var shouldHaveBorder = expectedBoundaries.Contains(row);

                            if (hasBorder != shouldHaveBorder)
                            {
                                if (shouldHaveBorder)
                                {
                                    return false.Label($"Row {row} should have thick bottom border (date: {dates[row - dataStartRow]}) but doesn't");
                                }
                                else
                                {
                                    return false.Label($"Row {row} should NOT have thick bottom border (date: {dates[row - dataStartRow]}) but does");
                                }
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Date group border application failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }
    }
}

