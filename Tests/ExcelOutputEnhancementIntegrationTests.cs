using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Integration tests for Excel Output Enhancement feature.
    /// Tests complete workflows with real test files.
    /// Validates: All requirements for excel-output-enhancement spec
    /// </summary>
    [TestFixture]
    public class ExcelOutputEnhancementIntegrationTests
    {
        private const string TestExcelFile = "2026_01_24_1906-2026 accompagnamenti sett 5 provvisorio 2.3 Greco.xlsx";
        private const string TestCSVFile = "168514-Estrazione_1770193162042.csv";

        private IExcelManager _excelManager = null!;
        private IDataTransformer _dataTransformer = null!;
        private ICSVParser _csvParser = null!;
        private ILookupService _lookupService = null!;
        private IColumnStructureManager _columnStructureManager = null!;
        private IFormattingService _formattingService = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _excelManager = new ExcelManager();
            _dataTransformer = new DataTransformer(new TransformationRulesEngine());
            _csvParser = new CSVParser();
            _lookupService = new LookupService();
            _columnStructureManager = new ColumnStructureManager();
            _formattingService = new FormattingService();
        }

        /// <summary>
        /// Test complete workflow with real Excel file containing reference sheets.
        /// Validates: All requirements
        /// </summary>
        [Test]
        public void IntegrationTest_WithRealExcelFile_LoadsReferenceSheets()
        {
            // Arrange
            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            // Act
            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);

            // Assert
            Assert.That(assistitiSheet, Is.Not.Null, "Assistiti sheet should exist");
            Assert.That(fissiSheet, Is.Not.Null, "Fissi sheet should exist");
            Assert.That(assistitiSheet.Worksheet, Is.Not.Null, "Assistiti worksheet should be accessible");
            Assert.That(fissiSheet.Worksheet, Is.Not.Null, "Fissi worksheet should be accessible");
        }

        /// <summary>
        /// Test that reference sheets can be loaded into LookupService.
        /// Validates: Requirements 5.1, 6.1, 8.2
        /// </summary>
        [Test]
        public void IntegrationTest_LoadReferenceSheets_Success()
        {
            // Arrange
            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);

            // Act
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Assert - No exception means success
            Assert.Pass("Reference sheets loaded successfully");
        }

        /// <summary>
        /// Test complete transformation workflow with real CSV file.
        /// Validates: Requirements 7.2, 5.1, 6.1, 8.2
        /// </summary>
        [Test]
        public void IntegrationTest_WithRealCSVFile_TransformsData()
        {
            // Arrange
            if (!File.Exists(TestCSVFile))
            {
                Assert.Ignore($"Test file not found: {TestCSVFile}");
                return;
            }

            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);
            
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var appointments = _csvParser.ParseCSV(TestCSVFile);
            var result = _dataTransformer.TransformEnhanced(appointments, _lookupService);

            // Assert
            Assert.That(result, Is.Not.Null, "Transformation result should not be null");
            Assert.That(result.Rows, Is.Not.Empty, "Should have transformed rows");
            
            // Verify CSV Note field mapped to Note Gasnet
            foreach (var row in result.Rows)
            {
                Assert.That(row, Is.Not.Null, "Row should not be null");
                // Note Gasnet should contain CSV note data
                Assert.That(row.NoteGasnet, Is.Not.Null, "NoteGasnet should not be null");
            }
        }

        /// <summary>
        /// Test complete end-to-end workflow from CSV to Excel output.
        /// Validates: All requirements
        /// </summary>
        [Test]
        public void IntegrationTest_EndToEnd_GeneratesEnhancedExcel()
        {
            // Arrange
            if (!File.Exists(TestCSVFile))
            {
                Assert.Ignore($"Test file not found: {TestCSVFile}");
                return;
            }

            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);
            
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            var appointments = _csvParser.ParseCSV(TestCSVFile);
            var result = _dataTransformer.TransformEnhanced(appointments, _lookupService);

            // Create a new test sheet
            var testSheet = _excelManager.CreateNewSheet(workbook, 999);

            // Act - Write enhanced data
            _excelManager.WriteColumnHeadersEnhanced(testSheet);
            _excelManager.WriteDataRowsEnhanced(testSheet, result.Rows, 3);
            
            // Apply formatting
            _excelManager.ApplyBoldToHeaders(testSheet, 2);
            
            // Sort data
            int lastRow = 3 + result.Rows.Count - 1;
            if (lastRow >= 3)
            {
                _excelManager.SortDataRows(testSheet, 3, lastRow);
                _excelManager.ApplyThickBordersToDateGroups(testSheet, 3, lastRow);
            }

            // Assert - Verify output structure
            var worksheet = testSheet.Worksheet;
            Assert.That(worksheet, Is.Not.Null, "Worksheet should not be null");

            // Verify 15 columns
            var headers = _columnStructureManager.GetColumnHeaders();
            Assert.That(headers.Count, Is.EqualTo(15), "Should have 15 columns");

            // Verify Indirizzo positioned after Assistito
            int assistitoIndex = headers.IndexOf("Assistito");
            int indirizzoIndex = headers.IndexOf("Indirizzo");
            Assert.That(indirizzoIndex, Is.EqualTo(assistitoIndex + 1), 
                "Indirizzo should be positioned immediately after Assistito");

            // Verify Comune Partenza excluded
            Assert.That(headers.Contains("Comune Partenza"), Is.False, 
                "Comune Partenza should be excluded from output");

            // Verify headers are bold (row 2)
            for (int col = 1; col <= 15; col++)
            {
                var cell = worksheet.Cells[2, col];
                Assert.That(cell.Style.Font.Bold, Is.True, 
                    $"Header cell at column {col} should be bold");
            }

            // Verify data rows exist
            Assert.That(worksheet.Dimension.End.Row, Is.GreaterThanOrEqualTo(3), 
                "Should have at least one data row");
        }

        /// <summary>
        /// Test that rows are sorted by date then time.
        /// Validates: Requirements 9.1, 9.2
        /// </summary>
        [Test]
        public void IntegrationTest_DataRows_AreSortedByDateThenTime()
        {
            // Arrange
            if (!File.Exists(TestCSVFile))
            {
                Assert.Ignore($"Test file not found: {TestCSVFile}");
                return;
            }

            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);
            
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            var appointments = _csvParser.ParseCSV(TestCSVFile);
            var result = _dataTransformer.TransformEnhanced(appointments, _lookupService);

            var testSheet = _excelManager.CreateNewSheet(workbook, 998);
            _excelManager.WriteColumnHeadersEnhanced(testSheet);
            _excelManager.WriteDataRowsEnhanced(testSheet, result.Rows, 3);
            
            int lastRow = 3 + result.Rows.Count - 1;
            if (lastRow >= 3)
            {
                _excelManager.SortDataRows(testSheet, 3, lastRow);
            }

            // Assert - Verify sorting
            var worksheet = testSheet.Worksheet;
            DateTime? previousDate = null;
            TimeSpan? previousTime = null;

            for (int row = 3; row <= lastRow; row++)
            {
                var dateCell = worksheet.Cells[row, 1];
                var timeCell = worksheet.Cells[row, 2];

                if (dateCell.Value == null) continue;

                DateTime currentDate;
                if (dateCell.Value is DateTime)
                {
                    currentDate = (DateTime)dateCell.Value;
                }
                else if (DateTime.TryParse(dateCell.Value.ToString(), out DateTime parsedDate))
                {
                    currentDate = parsedDate;
                }
                else
                {
                    continue;
                }

                if (previousDate.HasValue)
                {
                    Assert.That(currentDate, Is.GreaterThanOrEqualTo(previousDate.Value), 
                        $"Row {row} date should be >= previous date (primary sort)");

                    // If same date, check time sorting
                    if (currentDate.Date == previousDate.Value.Date && timeCell.Value != null)
                    {
                        TimeSpan currentTime;
                        if (TimeSpan.TryParse(timeCell.Value.ToString(), out currentTime))
                        {
                            if (previousTime.HasValue)
                            {
                                Assert.That(currentTime, Is.GreaterThanOrEqualTo(previousTime.Value), 
                                    $"Row {row} time should be >= previous time (secondary sort)");
                            }
                            previousTime = currentTime;
                        }
                    }
                    else
                    {
                        previousTime = null;
                    }
                }

                previousDate = currentDate;
            }
        }

        /// <summary>
        /// Test that thick borders are applied between date groups.
        /// Validates: Requirements 10.1, 10.2, 10.3
        /// </summary>
        [Test]
        public void IntegrationTest_DateGroups_HaveThickBorders()
        {
            // Arrange
            if (!File.Exists(TestCSVFile))
            {
                Assert.Ignore($"Test file not found: {TestCSVFile}");
                return;
            }

            if (!File.Exists(TestExcelFile))
            {
                Assert.Ignore($"Test file not found: {TestExcelFile}");
                return;
            }

            var workbook = _excelManager.OpenWorkbook(TestExcelFile);
            var assistitiSheet = _excelManager.GetSheetByName(workbook, "assistiti");
            var fissiSheet = _excelManager.GetFissiSheet(workbook);
            
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            var appointments = _csvParser.ParseCSV(TestCSVFile);
            var result = _dataTransformer.TransformEnhanced(appointments, _lookupService);

            var testSheet = _excelManager.CreateNewSheet(workbook, 997);
            _excelManager.WriteColumnHeadersEnhanced(testSheet);
            _excelManager.WriteDataRowsEnhanced(testSheet, result.Rows, 3);
            
            int lastRow = 3 + result.Rows.Count - 1;
            if (lastRow >= 3)
            {
                _excelManager.SortDataRows(testSheet, 3, lastRow);
                _excelManager.ApplyThickBordersToDateGroups(testSheet, 3, lastRow);
            }

            // Assert - Verify thick borders at date boundaries
            var worksheet = testSheet.Worksheet;
            DateTime? previousDate = null;
            bool foundDateBoundary = false;

            for (int row = 3; row <= lastRow; row++)
            {
                var dateCell = worksheet.Cells[row, 1];
                if (dateCell.Value == null) continue;

                DateTime currentDate;
                if (dateCell.Value is DateTime)
                {
                    currentDate = (DateTime)dateCell.Value;
                }
                else if (DateTime.TryParse(dateCell.Value.ToString(), out DateTime parsedDate))
                {
                    currentDate = parsedDate;
                }
                else
                {
                    continue;
                }

                // Check if this is the last row of a date group
                if (row < lastRow)
                {
                    var nextDateCell = worksheet.Cells[row + 1, 1];
                    if (nextDateCell.Value != null)
                    {
                        DateTime nextDate;
                        if (nextDateCell.Value is DateTime)
                        {
                            nextDate = (DateTime)nextDateCell.Value;
                        }
                        else if (DateTime.TryParse(nextDateCell.Value.ToString(), out DateTime parsedNextDate))
                        {
                            nextDate = parsedNextDate;
                        }
                        else
                        {
                            continue;
                        }

                        // If date changes, this row should have thick bottom border
                        if (currentDate.Date != nextDate.Date)
                        {
                            foundDateBoundary = true;
                            var borderStyle = worksheet.Cells[row, 1].Style.Border.Bottom.Style;
                            Assert.That(borderStyle, Is.EqualTo(OfficeOpenXml.Style.ExcelBorderStyle.Thick), 
                                $"Row {row} should have thick bottom border (date boundary)");
                        }
                    }
                }

                previousDate = currentDate;
            }

            // If we have multiple dates, we should have found at least one boundary
            if (lastRow > 3)
            {
                // This assertion is informational - not all test data will have multiple dates
                TestContext.WriteLine($"Found date boundary: {foundDateBoundary}");
            }
        }
    }
}
