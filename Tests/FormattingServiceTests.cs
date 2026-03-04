using System;
using System.Collections.Generic;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for FormattingService class.
    /// Tests specific examples and edge cases.
    /// Validates: Requirements 4.1, 10.1, 10.2, 10.3
    /// </summary>
    [TestFixture]
    public class FormattingServiceTests
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
        /// Helper method to create a test sheet with specified rows and columns
        /// </summary>
        private Sheet CreateTestSheet(int rows, int columns)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");

            for (int row = 1; row <= rows; row++)
            {
                for (int col = 1; col <= columns; col++)
                {
                    worksheet.Cells[row, col].Value = $"R{row}C{col}";
                }
            }

            return new Sheet(worksheet);
        }

        /// <summary>
        /// Test bold formatting on header row
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void ApplyBoldHeaders_ShouldApplyBoldToAllHeaderCells()
        {
            // Arrange
            var sheet = CreateTestSheet(5, 10);
            var headerRow = 1;

            // Act
            _formattingService.ApplyBoldHeaders(sheet, headerRow);

            // Assert
            for (int col = 1; col <= 10; col++)
            {
                var cell = sheet.Worksheet.Cells[headerRow, col];
                Assert.That(cell.Style.Font.Bold, Is.True, $"Column {col} header should be bold");
            }
        }

        /// <summary>
        /// Test that bold formatting is only applied to header row, not data rows
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void ApplyBoldHeaders_ShouldNotApplyBoldToDataRows()
        {
            // Arrange
            var sheet = CreateTestSheet(5, 10);
            var headerRow = 1;

            // Act
            _formattingService.ApplyBoldHeaders(sheet, headerRow);

            // Assert - Data rows should not be bold
            for (int row = 2; row <= 5; row++)
            {
                for (int col = 1; col <= 10; col++)
                {
                    var cell = sheet.Worksheet.Cells[row, col];
                    Assert.That(cell.Style.Font.Bold, Is.False, $"Data cell at row {row}, col {col} should not be bold");
                }
            }
        }

        /// <summary>
        /// Test date group border with single-row groups
        /// Validates: Requirements 10.1, 10.2
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_WithSingleRowGroups_ShouldApplyBorderToEachRow()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            
            // Create data with different dates (each row is its own group)
            worksheet.Cells[1, 1].Value = "Data";
            worksheet.Cells[2, 1].Value = "01/01/2024";
            worksheet.Cells[3, 1].Value = "02/01/2024";
            worksheet.Cells[4, 1].Value = "03/01/2024";
            
            var sheet = new Sheet(worksheet);

            // Act
            _formattingService.ApplyDateGroupBorders(sheet, 2, 4, 1);

            // Assert - Each row should have a thick bottom border
            for (int row = 2; row <= 4; row++)
            {
                var cell = sheet.Worksheet.Cells[row, 1];
                Assert.That(cell.Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                    $"Row {row} should have thick bottom border");
            }
        }

        /// <summary>
        /// Test date group border with multi-row groups
        /// Validates: Requirements 10.1, 10.2
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_WithMultiRowGroups_ShouldApplyBorderToLastRowOfEachGroup()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            
            // Create data with multiple rows per date group
            worksheet.Cells[1, 1].Value = "Data";
            worksheet.Cells[2, 1].Value = "01/01/2024";  // Group 1
            worksheet.Cells[3, 1].Value = "01/01/2024";  // Group 1
            worksheet.Cells[4, 1].Value = "01/01/2024";  // Group 1 - should have border
            worksheet.Cells[5, 1].Value = "02/01/2024";  // Group 2
            worksheet.Cells[6, 1].Value = "02/01/2024";  // Group 2 - should have border
            worksheet.Cells[7, 1].Value = "03/01/2024";  // Group 3 - should have border
            
            var sheet = new Sheet(worksheet);

            // Act
            _formattingService.ApplyDateGroupBorders(sheet, 2, 7, 1);

            // Assert - Only last row of each group should have thick bottom border
            Assert.That(sheet.Worksheet.Cells[2, 1].Style.Border.Bottom.Style, Is.Not.EqualTo(ExcelBorderStyle.Thick),
                "Row 2 should NOT have thick bottom border (not last in group)");
            Assert.That(sheet.Worksheet.Cells[3, 1].Style.Border.Bottom.Style, Is.Not.EqualTo(ExcelBorderStyle.Thick),
                "Row 3 should NOT have thick bottom border (not last in group)");
            Assert.That(sheet.Worksheet.Cells[4, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                "Row 4 should have thick bottom border (last in group 1)");
            Assert.That(sheet.Worksheet.Cells[5, 1].Style.Border.Bottom.Style, Is.Not.EqualTo(ExcelBorderStyle.Thick),
                "Row 5 should NOT have thick bottom border (not last in group)");
            Assert.That(sheet.Worksheet.Cells[6, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                "Row 6 should have thick bottom border (last in group 2)");
            Assert.That(sheet.Worksheet.Cells[7, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                "Row 7 should have thick bottom border (last in group 3)");
        }

        /// <summary>
        /// Test border application across all columns
        /// Validates: Requirements 10.2
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_ShouldApplyBorderToAllColumns()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("TestSheet");
            
            // Create data with 5 columns
            for (int col = 1; col <= 5; col++)
            {
                worksheet.Cells[1, col].Value = $"Column{col}";
                worksheet.Cells[2, col].Value = "01/01/2024";
                worksheet.Cells[3, col].Value = "02/01/2024";
            }
            
            var sheet = new Sheet(worksheet);

            // Act
            _formattingService.ApplyDateGroupBorders(sheet, 2, 3, 1);

            // Assert - All columns in boundary rows should have thick bottom border
            for (int col = 1; col <= 5; col++)
            {
                Assert.That(sheet.Worksheet.Cells[2, col].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                    $"Row 2, Column {col} should have thick bottom border");
                Assert.That(sheet.Worksheet.Cells[3, col].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                    $"Row 3, Column {col} should have thick bottom border");
            }
        }

        /// <summary>
        /// Test border application failure handling (log and continue)
        /// Validates: Requirements 10.3
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_WithNullSheet_ShouldThrowArgumentNullException()
        {
            // Arrange
            Sheet? nullSheet = null;

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() =>
                _formattingService.ApplyDateGroupBorders(nullSheet!, 2, 5, 1));
        }

        /// <summary>
        /// Test bold formatting with null sheet should throw exception
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void ApplyBoldHeaders_WithNullSheet_ShouldThrowArgumentNullException()
        {
            // Arrange
            Sheet? nullSheet = null;

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() =>
                _formattingService.ApplyBoldHeaders(nullSheet!, 1));
        }

        /// <summary>
        /// Test bold formatting with invalid header row should throw exception
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void ApplyBoldHeaders_WithInvalidHeaderRow_ShouldThrowArgumentException()
        {
            // Arrange
            var sheet = CreateTestSheet(5, 10);

            // Act & Assert
            Assert.Throws<ArgumentException>(() =>
                _formattingService.ApplyBoldHeaders(sheet, 0));
        }

        /// <summary>
        /// Test date group borders with invalid row range should throw exception
        /// Validates: Requirements 10.1
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_WithInvalidRowRange_ShouldThrowArgumentException()
        {
            // Arrange
            var sheet = CreateTestSheet(5, 10);

            // Act & Assert
            Assert.Throws<ArgumentException>(() =>
                _formattingService.ApplyDateGroupBorders(sheet, 5, 2, 1));
        }

        /// <summary>
        /// Test date group borders with empty sheet should handle gracefully
        /// Validates: Requirements 10.3
        /// </summary>
        [Test]
        public void ApplyDateGroupBorders_WithEmptySheet_ShouldNotThrow()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("EmptySheet");
            var sheet = new Sheet(worksheet);

            // Act & Assert - Should not throw
            Assert.DoesNotThrow(() =>
                _formattingService.ApplyDateGroupBorders(sheet, 2, 5, 1));
        }

        /// <summary>
        /// Test bold formatting with empty sheet should handle gracefully
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void ApplyBoldHeaders_WithEmptySheet_ShouldNotThrow()
        {
            // Arrange
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("EmptySheet");
            var sheet = new Sheet(worksheet);

            // Act & Assert - Should not throw
            Assert.DoesNotThrow(() =>
                _formattingService.ApplyBoldHeaders(sheet, 1));
        }
    }
}
