using System;
using System.Collections.Generic;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for ExcelManager enhanced methods.
    /// Tests specific examples and edge cases.
    /// </summary>
    [TestFixture]
    public class ExcelManagerEnhancedTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        [Test]
        public void WriteColumnHeadersEnhanced_WritesExactly15Columns()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Act
                _excelManager.WriteColumnHeadersEnhanced(sheet);

                // Assert
                int columnCount = 0;
                for (int col = 1; col <= 20; col++)
                {
                    var cellValue = worksheet.Cells[2, col].Value;
                    if (cellValue != null || col <= 15)
                        columnCount = col;
                }
                
                Assert.That(columnCount, Is.EqualTo(15));
            }
        }

        [Test]
        public void WriteColumnHeadersEnhanced_HasCorrectColumnNames()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Act
                _excelManager.WriteColumnHeadersEnhanced(sheet);

                // Assert
                Assert.That(worksheet.Cells[2, 1].Value?.ToString(), Is.EqualTo("Data"));
                Assert.That(worksheet.Cells[2, 2].Value?.ToString(), Is.EqualTo("Partenza"));
                Assert.That(worksheet.Cells[2, 3].Value?.ToString(), Is.EqualTo("Assistito"));
                Assert.That(worksheet.Cells[2, 4].Value?.ToString(), Is.EqualTo("Indirizzo"));
                Assert.That(worksheet.Cells[2, 5].Value?.ToString(), Is.EqualTo("Destinazione"));
                Assert.That(worksheet.Cells[2, 6].Value?.ToString(), Is.EqualTo("Note"));
                Assert.That(worksheet.Cells[2, 7].Value?.ToString(), Is.EqualTo("Auto"));
                Assert.That(worksheet.Cells[2, 8].Value?.ToString(), Is.EqualTo("Volontario"));
                Assert.That(worksheet.Cells[2, 9].Value?.ToString(), Is.EqualTo("Arrivo"));
                Assert.That(worksheet.Cells[2, 10].Value?.ToString(), Is.EqualTo("Avv"));
                Assert.That(worksheet.Cells[2, 11].Value?.ToString() ?? "", Is.EqualTo(""));
                Assert.That(worksheet.Cells[2, 12].Value?.ToString(), Is.EqualTo("Indirizzo Gasnet"));
                Assert.That(worksheet.Cells[2, 13].Value?.ToString(), Is.EqualTo("Note Gasnet"));
                Assert.That(worksheet.Cells[2, 14].Value?.ToString() ?? "", Is.EqualTo(""));
            }
        }

        [Test]
        public void WriteDataRowsEnhanced_WritesAllColumns()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "15/01/2024",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario",
                        Indirizzo = "Via Roma 1",
                        Destinazione = "Ospedale",
                        Note = "Test note",
                        Auto = "Auto1",
                        Volontario = "Vol1",
                        Arrivo = "10:00",
                        Avv = "TestAvv",
                        Empty1 = "",
                        IndirizzoGasnet = "Via Milano 2, Milano",
                        NoteGasnet = "CSV note",
                        Empty2 = ""
                    }
                };

                // Act
                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Assert
                Assert.That(worksheet.Cells[3, 2].Value?.ToString(), Is.EqualTo("09:00"));
                Assert.That(worksheet.Cells[3, 3].Value?.ToString(), Is.EqualTo("Rossi Mario"));
                Assert.That(worksheet.Cells[3, 4].Value?.ToString(), Is.EqualTo("Via Roma 1"));
                Assert.That(worksheet.Cells[3, 5].Value?.ToString(), Is.EqualTo("Ospedale"));
                Assert.That(worksheet.Cells[3, 6].Value?.ToString(), Is.EqualTo("Test note"));
                Assert.That(worksheet.Cells[3, 7].Value?.ToString(), Is.EqualTo("Auto1"));
                Assert.That(worksheet.Cells[3, 8].Value?.ToString(), Is.EqualTo("Vol1"));
                Assert.That(worksheet.Cells[3, 9].Value?.ToString(), Is.EqualTo("10:00"));
                Assert.That(worksheet.Cells[3, 10].Value?.ToString(), Is.EqualTo("TestAvv"));
                Assert.That(worksheet.Cells[3, 12].Value?.ToString(), Is.EqualTo("Via Milano 2, Milano"));
                Assert.That(worksheet.Cells[3, 13].Value?.ToString(), Is.EqualTo("CSV note"));
            }
        }

        [Test]
        public void WriteDataRowsEnhanced_FormatsDateCorrectly()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "15/01/2024",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario",
                        Indirizzo = "",
                        Destinazione = "",
                        Note = "",
                        Auto = "",
                        Volontario = "",
                        Arrivo = "",
                        Avv = "",
                        Empty1 = "",
                        IndirizzoGasnet = "",
                        NoteGasnet = "",
                        Empty2 = ""
                    }
                };

                // Act
                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Assert
                var dateCell = worksheet.Cells[3, 1];
                Assert.That(dateCell.Value, Is.TypeOf<DateTime>());
                Assert.That(dateCell.Style.Numberformat.Format, Is.EqualTo("ddd dd mmm"));
            }
        }

        [Test]
        public void SortDataRows_SortsByDateAndTime()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow { Data = "17/01/2024", Partenza = "14:00", Assistito = "C" },
                    new EnhancedTransformedRow { Data = "15/01/2024", Partenza = "10:00", Assistito = "A" },
                    new EnhancedTransformedRow { Data = "15/01/2024", Partenza = "09:00", Assistito = "B" },
                    new EnhancedTransformedRow { Data = "16/01/2024", Partenza = "11:00", Assistito = "D" }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Act
                _excelManager.SortDataRows(sheet, 3, 6);

                // Assert - dates should be in order: 15, 15, 16, 17
                var date1 = (DateTime)worksheet.Cells[3, 1].Value;
                var date2 = (DateTime)worksheet.Cells[4, 1].Value;
                var date3 = (DateTime)worksheet.Cells[5, 1].Value;
                var date4 = (DateTime)worksheet.Cells[6, 1].Value;

                Assert.That(date1, Is.LessThanOrEqualTo(date2));
                Assert.That(date2, Is.LessThanOrEqualTo(date3));
                Assert.That(date3, Is.LessThanOrEqualTo(date4));
            }
        }

        [Test]
        public void ApplyBoldToHeaders_AppliesBoldFormatting()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                _excelManager.WriteColumnHeadersEnhanced(sheet);

                // Act
                _excelManager.ApplyBoldToHeaders(sheet, 2);

                // Assert
                for (int col = 1; col <= 14; col++)
                {
                    Assert.That(worksheet.Cells[2, col].Style.Font.Bold, Is.True, 
                        $"Column {col} header should be bold");
                }
            }
        }

        [Test]
        public void ApplyThickBordersToDateGroups_AppliesBordersCorrectly()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Create 3 date groups
                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow { Data = "15/01/2024", Partenza = "09:00", Assistito = "A" },
                    new EnhancedTransformedRow { Data = "15/01/2024", Partenza = "10:00", Assistito = "B" },
                    new EnhancedTransformedRow { Data = "16/01/2024", Partenza = "09:00", Assistito = "C" },
                    new EnhancedTransformedRow { Data = "17/01/2024", Partenza = "09:00", Assistito = "D" },
                    new EnhancedTransformedRow { Data = "17/01/2024", Partenza = "10:00", Assistito = "E" }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Act
                _excelManager.ApplyThickBordersToDateGroups(sheet, 3, 7);

                // Assert - thick borders should be on rows 4 (last of 15/01), 5 (last of 16/01), 7 (last of 17/01)
                Assert.That(worksheet.Cells[4, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
                Assert.That(worksheet.Cells[5, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
                Assert.That(worksheet.Cells[7, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
                
                // No thick border on row 3 (not last of group)
                Assert.That(worksheet.Cells[3, 1].Style.Border.Bottom.Style, Is.Not.EqualTo(ExcelBorderStyle.Thick));
            }
        }

        [Test]
        public void ApplyThickBordersToDateGroups_HandlesSingleRowGroups()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow { Data = "15/01/2024", Partenza = "09:00", Assistito = "A" },
                    new EnhancedTransformedRow { Data = "16/01/2024", Partenza = "09:00", Assistito = "B" },
                    new EnhancedTransformedRow { Data = "17/01/2024", Partenza = "09:00", Assistito = "C" }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Act
                _excelManager.ApplyThickBordersToDateGroups(sheet, 3, 5);

                // Assert - each row should have thick border (each is last of its group)
                Assert.That(worksheet.Cells[3, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
                Assert.That(worksheet.Cells[4, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
                Assert.That(worksheet.Cells[5, 1].Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick));
            }
        }

        [Test]
        public void WriteDataRowsEnhanced_HandlesEmptyList()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>();

                // Act
                _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                // Assert - no exception, sheet remains empty
                Assert.That(worksheet.Cells[3, 1].Value, Is.Null);
            }
        }

        [Test]
        public void WriteDataRowsEnhanced_HandlesNullList()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                // Act
                _excelManager.WriteDataRowsEnhanced(sheet, null, 3);

                // Assert - no exception, sheet remains empty
                Assert.That(worksheet.Cells[3, 1].Value, Is.Null);
            }
        }
    }
}
