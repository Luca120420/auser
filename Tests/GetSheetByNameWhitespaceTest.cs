using System;
using NUnit.Framework;
using AuserExcelTransformer.Services;
using OfficeOpenXml;
using ExcelWorkbookModel = AuserExcelTransformer.Models.ExcelWorkbook;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Tests to verify GetSheetByName handles whitespace correctly.
    /// Validates that searching for "laboratori" finds both "laboratori" and "laboratori " sheets.
    /// </summary>
    [TestFixture]
    public class GetSheetByNameWhitespaceTest
    {
        private IExcelManager _excelManager = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        [Test]
        public void GetSheetByName_FindsSheetWithTrailingSpace_WhenSearchingWithoutSpace()
        {
            // Arrange - Create workbook with sheet named "laboratori " (with trailing space)
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("laboratori ");
                worksheet.Cells[1, 1].Value = "Test";
                
                var workbook = new ExcelWorkbookModel(package);

                // Act - Search for "laboratori" (without space)
                var result = _excelManager.GetSheetByName(workbook, "laboratori");

                // Assert
                Assert.That(result, Is.Not.Null, "Should find sheet with trailing space when searching without space");
                Assert.That(result.Worksheet.Name, Is.EqualTo("laboratori "), "Should return the actual sheet with trailing space");
            }
        }

        [Test]
        public void GetSheetByName_FindsSheetWithoutSpace_WhenSearchingWithoutSpace()
        {
            // Arrange - Create workbook with sheet named "laboratori" (no space)
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("laboratori");
                worksheet.Cells[1, 1].Value = "Test";
                
                var workbook = new ExcelWorkbookModel(package);

                // Act - Search for "laboratori" (without space)
                var result = _excelManager.GetSheetByName(workbook, "laboratori");

                // Assert
                Assert.That(result, Is.Not.Null, "Should find sheet without space when searching without space");
                Assert.That(result.Worksheet.Name, Is.EqualTo("laboratori"), "Should return the actual sheet without space");
            }
        }

        [Test]
        public void GetSheetByName_FindsSheetWithTrailingSpace_WhenSearchingWithSpace()
        {
            // Arrange - Create workbook with sheet named "laboratori " (with trailing space)
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("laboratori ");
                worksheet.Cells[1, 1].Value = "Test";
                
                var workbook = new ExcelWorkbookModel(package);

                // Act - Search for "laboratori " (with space)
                var result = _excelManager.GetSheetByName(workbook, "laboratori ");

                // Assert
                Assert.That(result, Is.Not.Null, "Should find sheet with trailing space when searching with space");
                Assert.That(result.Worksheet.Name, Is.EqualTo("laboratori "), "Should return the actual sheet with trailing space");
            }
        }

        [Test]
        public void GetSheetByName_FindsSheetWithoutSpace_WhenSearchingWithSpace()
        {
            // Arrange - Create workbook with sheet named "laboratori" (no space)
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("laboratori");
                worksheet.Cells[1, 1].Value = "Test";
                
                var workbook = new ExcelWorkbookModel(package);

                // Act - Search for "laboratori " (with space)
                var result = _excelManager.GetSheetByName(workbook, "laboratori ");

                // Assert
                Assert.That(result, Is.Not.Null, "Should find sheet without space when searching with space");
                Assert.That(result.Worksheet.Name, Is.EqualTo("laboratori"), "Should return the actual sheet without space");
            }
        }

        [Test]
        public void GetSheetByName_IsCaseInsensitive()
        {
            // Arrange - Create workbook with sheet named "Laboratori " (capital L, with space)
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Laboratori ");
                worksheet.Cells[1, 1].Value = "Test";
                
                var workbook = new ExcelWorkbookModel(package);

                // Act - Search for "laboratori" (lowercase, no space)
                var result = _excelManager.GetSheetByName(workbook, "laboratori");

                // Assert
                Assert.That(result, Is.Not.Null, "Should find sheet with case-insensitive match");
                Assert.That(result.Worksheet.Name, Is.EqualTo("Laboratori "), "Should return the actual sheet");
            }
        }
    }
}
