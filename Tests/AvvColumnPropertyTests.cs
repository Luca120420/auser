using System;
using System.Collections.Generic;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for AVV column header position verification.
    /// Validates: Requirements 1.4, 1.5
    /// </summary>
    [TestFixture]
    public class AvvColumnHeaderPositionTests
    {
        private ColumnStructureManager _columnStructureManager = null!;
        private ExcelManager _excelManager = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _columnStructureManager = new ColumnStructureManager();
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Verify that "Avvisi" is at index 9 (column 10) in the column headers.
        /// Validates: Requirement 1.4
        /// </summary>
        [Test]
        public void GetColumnHeaders_AvvIsAtIndex9()
        {
            // Act
            var headers = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers[9], Is.EqualTo("Avvisi"),
                "Column at index 9 (position 10) should be 'Avvisi'");
        }

        /// <summary>
        /// Verify that "Arrivo" is at index 8 (column 9) to confirm ordering.
        /// Avv must come immediately after Arrivo.
        /// Validates: Requirement 1.4
        /// </summary>
        [Test]
        public void GetColumnHeaders_ArrivoIsAtIndex8_BeforeAvv()
        {
            // Act
            var headers = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers[8], Is.EqualTo("Arrivo"),
                "Column at index 8 (position 9) should be 'Arrivo'");
        }

        /// <summary>
        /// Verify that Avv immediately follows Arrivo in the column order.
        /// Validates: Requirement 1.4
        /// </summary>
        [Test]
        public void GetColumnHeaders_AvvImmediatelyFollowsArrivo()
        {
            // Act
            var headers = _columnStructureManager.GetColumnHeaders();
            int arrivoIndex = headers.IndexOf("Arrivo");
            int avvIndex = headers.IndexOf("Avvisi");

            // Assert
            Assert.That(avvIndex, Is.EqualTo(arrivoIndex + 1),
                "Avv should be positioned immediately after Arrivo");
        }

        /// <summary>
        /// Verify that WriteColumnHeadersEnhanced writes "Avvisi" in cell [2, 10].
        /// Validates: Requirement 1.5
        /// </summary>
        [Test]
        public void WriteColumnHeadersEnhanced_WritesAvvInCell2_10()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Act
                _excelManager.WriteColumnHeadersEnhanced(sheet);

                // Assert
                Assert.That(worksheet.Cells[2, 10].Value?.ToString(), Is.EqualTo("Avvisi"),
                    "Cell [2, 10] should contain 'Avvisi' as the column header");
            }
        }

        /// <summary>
        /// Verify that WriteColumnHeadersEnhanced writes "Arrivo" in cell [2, 9]
        /// confirming the correct ordering in the actual output sheet.
        /// Validates: Requirement 1.4, 1.5
        /// </summary>
        [Test]
        public void WriteColumnHeadersEnhanced_WritesArrivoInCell2_9()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Act
                _excelManager.WriteColumnHeadersEnhanced(sheet);

                // Assert
                Assert.That(worksheet.Cells[2, 9].Value?.ToString(), Is.EqualTo("Arrivo"),
                    "Cell [2, 9] should contain 'Arrivo' as the column header");
            }
        }
    }
}
