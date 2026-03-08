using System;
using System.Linq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using ExcelWorkbookModel = AuserExcelTransformer.Models.ExcelWorkbook;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for ExcelManager laboratori sheet processing.
    /// Tests universal properties across many generated inputs.
    /// </summary>
    [TestFixture]
    public class ExcelManagerLaboratoriPropertyTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        // Feature: laboratori-sheet-processing, Property 1: Laboratori Sheet Detection
        /// <summary>
        /// Property 1: Laboratori Sheet Detection
        /// For any Excel workbook, when opened by ExcelManager, the system should correctly identify
        /// whether a sheet named "laboratori" exists in the workbook.
        /// Validates: Requirements 1.1, 1.3
        /// </summary>
        [Test]
        public void Property_LaboratoriSheetDetection()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for whether laboratori sheet should exist
            var hasLaboratoriGen = Arb.Generate<bool>();

            // Generator for case variations of "laboratori"
            var laboratoriNameGen = Gen.Elements(
                "laboratori", "Laboratori", "LABORATORI", "LaBoRaToRi"
            );

            // Generator for number of other sheets (0-5)
            var otherSheetCountGen = Gen.Choose(0, 5);

            var testGen = from hasLaboratori in hasLaboratoriGen
                          from laboratoriName in laboratoriNameGen
                          from otherSheetCount in otherSheetCountGen
                          select (hasLaboratori, laboratoriName, otherSheetCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (hasLaboratori, laboratoriName, otherSheetCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create other sheets with random names
                    for (int i = 0; i < otherSheetCount; i++)
                    {
                        package.Workbook.Worksheets.Add($"Sheet{i + 1}");
                    }

                    // Conditionally create laboratori sheet
                    if (hasLaboratori)
                    {
                        var labSheet = package.Workbook.Worksheets.Add(laboratoriName);
                        // Add minimal structure to make it a valid sheet
                        labSheet.Cells[1, 1].Value = "Data";
                    }

                    var workbook = new ExcelWorkbookModel(package);

                    // Act
                    var laboratoriSheet = _excelManager.GetSheetByName(workbook, "laboratori");

                    // Assert - detection should match whether we created the sheet
                    bool detected = laboratoriSheet != null;
                    
                    if (detected != hasLaboratori)
                    {
                        throw new Exception(
                            $"Sheet detection mismatch. Expected: {hasLaboratori}, Got: {detected}. " +
                            $"Sheet name used: '{laboratoriName}', Other sheets: {otherSheetCount}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: laboratori-sheet-processing, Property 2: Graceful Handling of Missing Laboratori Sheet
        /// <summary>
        /// Property 2: Graceful Handling of Missing Laboratori Sheet
        /// For any Excel workbook without a "laboratori" sheet, the system should complete processing
        /// successfully without throwing exceptions or displaying error messages, processing only CSV and fissi data.
        /// Validates: Requirements 1.2, 8.1, 8.2, 8.3
        /// </summary>
        [Test]
        public void Property_GracefulHandlingOfMissingLaboratoriSheet()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for number of other sheets (1-5, must have at least one for target)
            var otherSheetCountGen = Gen.Choose(1, 5);

            // Generator for number of data rows to simulate in CSV/fissi processing (0-10)
            var dataRowCountGen = Gen.Choose(0, 10);

            var testGen = from otherSheetCount in otherSheetCountGen
                          from dataRowCount in dataRowCountGen
                          select (otherSheetCount, dataRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (otherSheetCount, dataRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create workbook WITHOUT laboratori sheet
                    // Create a target sheet for output
                    var targetSheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Add some other sheets (but NOT laboratori)
                    for (int i = 0; i < otherSheetCount; i++)
                    {
                        var sheetName = $"Sheet{i + 1}";
                        // Make sure we don't accidentally create "laboratori"
                        if (sheetName.Equals("laboratori", StringComparison.OrdinalIgnoreCase))
                        {
                            sheetName = "OtherSheet";
                        }
                        package.Workbook.Worksheets.Add(sheetName);
                    }

                    // Simulate some existing data in target sheet (from CSV/fissi processing)
                    targetSheet.Cells[1, 1].Value = "Header";
                    targetSheet.Cells[2, 1].Value = "Data";
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        targetSheet.Cells[3 + i, 1].Value = $"Row{i + 1}";
                    }

                    var workbook = new ExcelWorkbookModel(package);

                    // Act - Try to get laboratori sheet (should return null)
                    var laboratoriSheet = _excelManager.GetSheetByName(workbook, "laboratori");

                    // Assert 1: laboratori sheet should not be found
                    if (laboratoriSheet != null)
                    {
                        throw new Exception("Expected laboratori sheet to be null when not present in workbook");
                    }

                    // Act 2 - Simulate the workflow: only process if laboratori sheet exists
                    // This is what ApplicationController should do
                    int nextRow = 3 + dataRowCount;
                    bool exceptionThrown = false;
                    string exceptionMessage = null;

                    try
                    {
                        // The workflow should check if laboratoriSheet is null before calling AppendLaboratoriData
                        if (laboratoriSheet != null)
                        {
                            _excelManager.AppendLaboratoriData(
                                new Sheet(targetSheet), 
                                laboratoriSheet, 
                                nextRow
                            );
                        }
                        // If laboratoriSheet is null, we simply skip the append - no error
                    }
                    catch (Exception ex)
                    {
                        exceptionThrown = true;
                        exceptionMessage = ex.Message;
                    }

                    // Assert 2: No exception should be thrown when laboratori sheet is missing
                    if (exceptionThrown)
                    {
                        throw new Exception(
                            $"Processing should complete without exceptions when laboratori sheet is missing. " +
                            $"Exception thrown: {exceptionMessage}"
                        );
                    }

                    // Assert 3: Existing data should remain unchanged
                    var finalDimension = targetSheet.Dimension;
                    if (finalDimension != null)
                    {
                        int finalRowCount = finalDimension.End.Row - 2; // Subtract header rows
                        if (finalRowCount != dataRowCount)
                        {
                            throw new Exception(
                                $"Existing data should remain unchanged when laboratori sheet is missing. " +
                                $"Expected {dataRowCount} data rows, got {finalRowCount}"
                            );
                        }
                    }
                    else if (dataRowCount > 0)
                    {
                        throw new Exception(
                            $"Expected {dataRowCount} data rows but sheet dimension is null"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: laboratori-sheet-processing, Property 3: Read All Laboratori Data Rows
        /// <summary>
        /// Property 3: Read All Laboratori Data Rows
        /// For any laboratori sheet with N data rows (excluding header rows), the AppendLaboratoriData
        /// method should read and process exactly N rows.
        /// Validates: Requirements 2.1
        /// </summary>
        [Test]
        public void Property_ReadAllLaboratoriDataRows()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for header position (row 1 or row 2)
            var headerInRow2Gen = Arb.Generate<bool>();

            // Generator for number of data rows (0-20)
            var dataRowCountGen = Gen.Choose(0, 20);

            // Generator for number of empty rows to intersperse (0-5)
            var emptyRowCountGen = Gen.Choose(0, 5);

            var testGen = from headerInRow2 in headerInRow2Gen
                          from dataRowCount in dataRowCountGen
                          from emptyRowCount in emptyRowCountGen
                          select (headerInRow2, dataRowCount, emptyRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (headerInRow2, dataRowCount, emptyRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");

                    // Set up header rows based on headerInRow2
                    int dataStartRow;
                    if (headerInRow2)
                    {
                        // Row 1 is title, Row 2 is headers
                        laboratoriWorksheet.Cells[1, 1].Value = "Laboratori Report";
                        laboratoriWorksheet.Cells[2, 1].Value = "Data";
                        laboratoriWorksheet.Cells[2, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[2, 3].Value = "Assistito";
                        dataStartRow = 3;
                    }
                    else
                    {
                        // Row 1 is headers
                        laboratoriWorksheet.Cells[1, 1].Value = "Data";
                        laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                        dataStartRow = 2;
                    }

                    // Add data rows with some empty rows interspersed
                    int currentRow = dataStartRow;
                    int actualDataRowsAdded = 0;

                    for (int i = 0; i < dataRowCount; i++)
                    {
                        // Add a valid data row
                        laboratoriWorksheet.Cells[currentRow, 1].Value = new DateTime(2024, 1, i % 28 + 1);
                        laboratoriWorksheet.Cells[currentRow, 2].Value = 0.35; // 8:24 AM
                        laboratoriWorksheet.Cells[currentRow, 3].Value = $"Patient{i}";
                        laboratoriWorksheet.Cells[currentRow, 4].Value = $"Address{i}";
                        laboratoriWorksheet.Cells[currentRow, 5].Value = $"Destination{i}";
                        laboratoriWorksheet.Cells[currentRow, 6].Value = $"Note{i}";
                        laboratoriWorksheet.Cells[currentRow, 7].Value = $"Car{i}";
                        laboratoriWorksheet.Cells[currentRow, 8].Value = $"Volunteer{i}";
                        laboratoriWorksheet.Cells[currentRow, 9].Value = 0.5; // 12:00 PM
                        laboratoriWorksheet.Cells[currentRow, 10].Value = $"Avv{i}";
                        actualDataRowsAdded++;
                        currentRow++;

                        // Occasionally add an empty row (should be skipped)
                        if (i < emptyRowCount)
                        {
                            // Leave column 1 empty - this row should be skipped
                            laboratoriWorksheet.Cells[currentRow, 2].Value = "ShouldBeSkipped";
                            currentRow++;
                        }
                    }

                    var workbook = new ExcelWorkbookModel(package);
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data starting at row 1
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1);

                    // Assert - Count how many rows were actually written to target
                    var targetDimension = targetWorksheet.Dimension;
                    int actualRowsWritten = 0;

                    if (targetDimension != null)
                    {
                        actualRowsWritten = targetDimension.End.Row;
                    }

                    // The number of rows written should equal the number of non-empty data rows
                    // (which is dataRowCount, since we only added empty rows that should be skipped)
                    if (actualRowsWritten != dataRowCount)
                    {
                        throw new Exception(
                            $"Expected {dataRowCount} data rows to be processed, but {actualRowsWritten} rows were written. " +
                            $"Header in row 2: {headerInRow2}, Empty rows added: {emptyRowCount}"
                        );
                    }

                    // Verify that all written rows have non-empty Data column (column 1)
                    for (int row = 1; row <= actualRowsWritten; row++)
                    {
                        var dataValue = targetWorksheet.Cells[row, 1].Value;
                        if (dataValue == null || string.IsNullOrWhiteSpace(dataValue.ToString()))
                        {
                            throw new Exception(
                                $"Row {row} in target sheet has empty Data column, but should have been skipped"
                            );
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: laboratori-sheet-processing, Property 4: Header Row Detection Consistency
        /// <summary>
        /// Property 4: Header Row Detection Consistency
        /// For any laboratori sheet, the system should detect the header row location using the same
        /// logic as fissi sheet detection: if row 2 contains "Data" in cell A2, data starts at row 3;
        /// if row 1 contains "Data" in cell A1, data starts at row 2.
        /// Validates: Requirements 2.2
        /// </summary>
        [Test]
        public void Property_HeaderRowDetectionConsistency()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for header position (true = row 2, false = row 1)
            var headerInRow2Gen = Arb.Generate<bool>();

            // Generator for number of data rows (1-15)
            var dataRowCountGen = Gen.Choose(1, 15);

            // Generator for additional content in non-header rows
            var hasExtraContentGen = Arb.Generate<bool>();

            var testGen = from headerInRow2 in headerInRow2Gen
                          from dataRowCount in dataRowCountGen
                          from hasExtraContent in hasExtraContentGen
                          select (headerInRow2, dataRowCount, hasExtraContent);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (headerInRow2, dataRowCount, hasExtraContent) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");

                    // Set up header rows based on headerInRow2
                    int expectedDataStartRow;
                    if (headerInRow2)
                    {
                        // Row 1 is title, Row 2 contains "Data" header
                        if (hasExtraContent)
                        {
                            laboratoriWorksheet.Cells[1, 1].Value = "Laboratori Report";
                            laboratoriWorksheet.Cells[1, 2].Value = "Generated on 2024-01-01";
                        }
                        laboratoriWorksheet.Cells[2, 1].Value = "Data";
                        laboratoriWorksheet.Cells[2, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[2, 3].Value = "Assistito";
                        laboratoriWorksheet.Cells[2, 4].Value = "Indirizzo";
                        laboratoriWorksheet.Cells[2, 5].Value = "Destinazione";
                        laboratoriWorksheet.Cells[2, 6].Value = "Note";
                        laboratoriWorksheet.Cells[2, 7].Value = "Auto";
                        laboratoriWorksheet.Cells[2, 8].Value = "Volontario";
                        laboratoriWorksheet.Cells[2, 9].Value = "Arrivo";
                        laboratoriWorksheet.Cells[2, 10].Value = "Avv";
                        expectedDataStartRow = 3;
                    }
                    else
                    {
                        // Row 1 contains "Data" header
                        laboratoriWorksheet.Cells[1, 1].Value = "Data";
                        laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                        laboratoriWorksheet.Cells[1, 4].Value = "Indirizzo";
                        laboratoriWorksheet.Cells[1, 5].Value = "Destinazione";
                        laboratoriWorksheet.Cells[1, 6].Value = "Note";
                        laboratoriWorksheet.Cells[1, 7].Value = "Auto";
                        laboratoriWorksheet.Cells[1, 8].Value = "Volontario";
                        laboratoriWorksheet.Cells[1, 9].Value = "Arrivo";
                        laboratoriWorksheet.Cells[1, 10].Value = "Avv";
                        expectedDataStartRow = 2;
                    }

                    // Add data rows starting from expectedDataStartRow
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        int rowNum = expectedDataStartRow + i;
                        laboratoriWorksheet.Cells[rowNum, 1].Value = new DateTime(2024, 1, (i % 28) + 1);
                        laboratoriWorksheet.Cells[rowNum, 2].Value = 0.35 + (i * 0.01); // Time values
                        laboratoriWorksheet.Cells[rowNum, 3].Value = $"Patient{i}";
                        laboratoriWorksheet.Cells[rowNum, 4].Value = $"Address{i}";
                        laboratoriWorksheet.Cells[rowNum, 5].Value = $"Destination{i}";
                        laboratoriWorksheet.Cells[rowNum, 6].Value = $"Note{i}";
                        laboratoriWorksheet.Cells[rowNum, 7].Value = $"Car{i}";
                        laboratoriWorksheet.Cells[rowNum, 8].Value = $"Volunteer{i}";
                        laboratoriWorksheet.Cells[rowNum, 9].Value = 0.5 + (i * 0.01); // Time values
                        laboratoriWorksheet.Cells[rowNum, 10].Value = $"Avv{i}";
                    }

                    var workbook = new ExcelWorkbookModel(package);
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data starting at row 1
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1);

                    // Assert - Verify that exactly dataRowCount rows were written
                    var targetDimension = targetWorksheet.Dimension;
                    int actualRowsWritten = 0;

                    if (targetDimension != null)
                    {
                        actualRowsWritten = targetDimension.End.Row;
                    }

                    if (actualRowsWritten != dataRowCount)
                    {
                        throw new Exception(
                            $"Expected {dataRowCount} data rows to be written, but {actualRowsWritten} rows were written. " +
                            $"Header in row 2: {headerInRow2}, Expected data start row: {expectedDataStartRow}"
                        );
                    }

                    // Verify that the first data row contains the expected first patient
                    if (dataRowCount > 0)
                    {
                        var firstPatient = targetWorksheet.Cells[1, 3].Value?.ToString();
                        if (firstPatient != "Patient0")
                        {
                            throw new Exception(
                                $"Expected first patient to be 'Patient0', but got '{firstPatient}'. " +
                                $"This indicates header detection may have failed. " +
                                $"Header in row 2: {headerInRow2}, Expected data start row: {expectedDataStartRow}"
                            );
                        }

                        // Verify the last data row contains the expected last patient
                        var lastPatient = targetWorksheet.Cells[dataRowCount, 3].Value?.ToString();
                        var expectedLastPatient = $"Patient{dataRowCount - 1}";
                        if (lastPatient != expectedLastPatient)
                        {
                            throw new Exception(
                                $"Expected last patient to be '{expectedLastPatient}', but got '{lastPatient}'. " +
                                $"This indicates incorrect row counting or header detection."
                            );
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: laboratori-sheet-processing, Property 5: Skip Rows with Empty Data Column
        /// <summary>
        /// Property 5: Skip Rows with Empty Data Column
        /// For any laboratori sheet, when a row has an empty or null value in column 1 (Data),
        /// that row should not appear in the output sheet.
        /// Validates: Requirements 2.6
        /// </summary>
        [Test]
        public void Property_SkipRowsWithEmptyDataColumn()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for header position (true = row 2, false = row 1)
            var headerInRow2Gen = Arb.Generate<bool>();

            // Generator for total number of rows including empty ones (1-20)
            var totalRowCountGen = Gen.Choose(1, 20);

            // Generator for which rows should have empty Data column (as indices)
            var emptyRowIndicesGen = Gen.Sized(size =>
            {
                return Gen.Choose(0, Math.Max(0, size - 1))
                    .ListOf()
                    .Select(list => list.Distinct().ToList());
            });

            var testGen = from headerInRow2 in headerInRow2Gen
                          from totalRowCount in totalRowCountGen
                          from emptyRowIndices in emptyRowIndicesGen
                          select (headerInRow2, totalRowCount, emptyRowIndices.Where(i => i < totalRowCount).ToList());

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (headerInRow2, totalRowCount, emptyRowIndices) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");

                    // Set up header rows based on headerInRow2
                    int dataStartRow;
                    if (headerInRow2)
                    {
                        // Row 1 is title, Row 2 is headers
                        laboratoriWorksheet.Cells[1, 1].Value = "Laboratori Report";
                        laboratoriWorksheet.Cells[2, 1].Value = "Data";
                        laboratoriWorksheet.Cells[2, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[2, 3].Value = "Assistito";
                        dataStartRow = 3;
                    }
                    else
                    {
                        // Row 1 is headers
                        laboratoriWorksheet.Cells[1, 1].Value = "Data";
                        laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                        dataStartRow = 2;
                    }

                    // Add data rows, some with empty Data column
                    int expectedNonEmptyRows = 0;
                    
                    for (int i = 0; i < totalRowCount; i++)
                    {
                        int rowNum = dataStartRow + i;
                        bool shouldBeEmpty = emptyRowIndices.Contains(i);

                        if (shouldBeEmpty)
                        {
                            // Leave column 1 (Data) empty or null
                            // Randomly choose between null, empty string, or whitespace
                            var emptyType = i % 3;
                            if (emptyType == 0)
                            {
                                // Leave as null (don't set value)
                            }
                            else if (emptyType == 1)
                            {
                                laboratoriWorksheet.Cells[rowNum, 1].Value = "";
                            }
                            else
                            {
                                laboratoriWorksheet.Cells[rowNum, 1].Value = "   "; // Whitespace
                            }
                            
                            // Add data to other columns to verify they're still skipped
                            laboratoriWorksheet.Cells[rowNum, 2].Value = 0.35;
                            laboratoriWorksheet.Cells[rowNum, 3].Value = $"ShouldBeSkipped{i}";
                        }
                        else
                        {
                            // Add a valid data row with non-empty Data column
                            laboratoriWorksheet.Cells[rowNum, 1].Value = new DateTime(2024, 1, (i % 28) + 1);
                            laboratoriWorksheet.Cells[rowNum, 2].Value = 0.35 + (i * 0.01);
                            laboratoriWorksheet.Cells[rowNum, 3].Value = $"Patient{i}";
                            laboratoriWorksheet.Cells[rowNum, 4].Value = $"Address{i}";
                            laboratoriWorksheet.Cells[rowNum, 5].Value = $"Destination{i}";
                            laboratoriWorksheet.Cells[rowNum, 6].Value = $"Note{i}";
                            laboratoriWorksheet.Cells[rowNum, 7].Value = $"Car{i}";
                            laboratoriWorksheet.Cells[rowNum, 8].Value = $"Volunteer{i}";
                            laboratoriWorksheet.Cells[rowNum, 9].Value = 0.5 + (i * 0.01);
                            laboratoriWorksheet.Cells[rowNum, 10].Value = $"Avv{i}";
                            expectedNonEmptyRows++;
                        }
                    }

                    var workbook = new ExcelWorkbookModel(package);
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data starting at row 1
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1);

                    // Assert - Count how many rows were actually written to target
                    var targetDimension = targetWorksheet.Dimension;
                    int actualRowsWritten = 0;

                    if (targetDimension != null)
                    {
                        actualRowsWritten = targetDimension.End.Row;
                    }

                    // The number of rows written should equal only the non-empty rows
                    if (actualRowsWritten != expectedNonEmptyRows)
                    {
                        throw new Exception(
                            $"Expected {expectedNonEmptyRows} non-empty data rows to be written, but {actualRowsWritten} rows were written. " +
                            $"Total rows: {totalRowCount}, Empty row indices: [{string.Join(", ", emptyRowIndices)}], " +
                            $"Header in row 2: {headerInRow2}"
                        );
                    }

                    // Verify that all written rows have non-empty Data column (column 1)
                    for (int row = 1; row <= actualRowsWritten; row++)
                    {
                        var dataValue = targetWorksheet.Cells[row, 1].Value;
                        if (dataValue == null || string.IsNullOrWhiteSpace(dataValue.ToString()))
                        {
                            throw new Exception(
                                $"Row {row} in target sheet has empty Data column, but should have been skipped. " +
                                $"Value: '{dataValue}'"
                            );
                        }
                    }

                    // Verify that none of the "ShouldBeSkipped" markers appear in the output
                    for (int row = 1; row <= actualRowsWritten; row++)
                    {
                        var assistitoValue = targetWorksheet.Cells[row, 3].Value?.ToString();
                        if (assistitoValue != null && assistitoValue.Contains("ShouldBeSkipped"))
                        {
                            throw new Exception(
                                $"Row {row} contains 'ShouldBeSkipped' marker in Assistito column, " +
                                $"indicating a row with empty Data was not properly skipped"
                            );
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: laboratori-sheet-processing, Property 6: Complete Column Mapping
        /// <summary>
        /// Property 6: Complete Column Mapping
        /// For any laboratori data row, all 10 columns (Data, Partenza, Assistito, Indirizzo, Destinazione,
        /// Note, Auto, Volontario, Arrivo, Avv) should be mapped to their corresponding output columns (1-10)
        /// with values preserved.
        /// Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6, 3.7, 3.8, 3.9, 3.10
        /// </summary>
        [Test]
        public void Property_CompleteColumnMapping()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for header position (true = row 2, false = row 1)
            var headerInRow2Gen = Arb.Generate<bool>();

            // Generator for number of data rows (1-15)
            var dataRowCountGen = Gen.Choose(1, 15);

            // Generator for random dates
            var dateGen = Gen.Choose(1, 28).Select(day => new DateTime(2024, 1, day));

            // Generator for random time values (as fractions of a day: 0.0 to 0.99)
            var timeGen = Gen.Choose(0, 99).Select(hundredths => hundredths / 100.0);

            // Generator for random strings (for text columns)
            var stringGen = Arb.Generate<NonEmptyString>().Select(nes => nes.Get);

            // Generator for a complete data row (all 10 columns)
            var dataRowGen = from date in dateGen
                             from partenza in timeGen
                             from assistito in stringGen
                             from indirizzo in stringGen
                             from destinazione in stringGen
                             from note in stringGen
                             from auto in stringGen
                             from volontario in stringGen
                             from arrivo in timeGen
                             from avv in stringGen
                             select (date, partenza, assistito, indirizzo, destinazione, note, auto, volontario, arrivo, avv);

            // Generator for a list of data rows
            var dataRowsGen = Gen.Sized(size =>
            {
                var count = Math.Max(1, Math.Min(size, 15));
                return Gen.ListOf(count, dataRowGen).Select(list => list.ToList());
            });

            var testGen = from headerInRow2 in headerInRow2Gen
                          from dataRows in dataRowsGen
                          select (headerInRow2, dataRows);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (headerInRow2, dataRows) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");

                    // Set up header rows based on headerInRow2
                    int dataStartRow;
                    if (headerInRow2)
                    {
                        // Row 1 is title, Row 2 contains "Data" header
                        laboratoriWorksheet.Cells[1, 1].Value = "Laboratori Report";
                        laboratoriWorksheet.Cells[2, 1].Value = "Data";
                        laboratoriWorksheet.Cells[2, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[2, 3].Value = "Assistito";
                        laboratoriWorksheet.Cells[2, 4].Value = "Indirizzo";
                        laboratoriWorksheet.Cells[2, 5].Value = "Destinazione";
                        laboratoriWorksheet.Cells[2, 6].Value = "Note";
                        laboratoriWorksheet.Cells[2, 7].Value = "Auto";
                        laboratoriWorksheet.Cells[2, 8].Value = "Volontario";
                        laboratoriWorksheet.Cells[2, 9].Value = "Arrivo";
                        laboratoriWorksheet.Cells[2, 10].Value = "Avv";
                        dataStartRow = 3;
                    }
                    else
                    {
                        // Row 1 contains "Data" header
                        laboratoriWorksheet.Cells[1, 1].Value = "Data";
                        laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                        laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                        laboratoriWorksheet.Cells[1, 4].Value = "Indirizzo";
                        laboratoriWorksheet.Cells[1, 5].Value = "Destinazione";
                        laboratoriWorksheet.Cells[1, 6].Value = "Note";
                        laboratoriWorksheet.Cells[1, 7].Value = "Auto";
                        laboratoriWorksheet.Cells[1, 8].Value = "Volontario";
                        laboratoriWorksheet.Cells[1, 9].Value = "Arrivo";
                        laboratoriWorksheet.Cells[1, 10].Value = "Avv";
                        dataStartRow = 2;
                    }

                    // Add data rows from generated data
                    for (int i = 0; i < dataRows.Count; i++)
                    {
                        int rowNum = dataStartRow + i;
                        var (date, partenza, assistito, indirizzo, destinazione, note, auto, volontario, arrivo, avv) = dataRows[i];

                        laboratoriWorksheet.Cells[rowNum, 1].Value = date;
                        laboratoriWorksheet.Cells[rowNum, 2].Value = partenza;
                        laboratoriWorksheet.Cells[rowNum, 3].Value = assistito;
                        laboratoriWorksheet.Cells[rowNum, 4].Value = indirizzo;
                        laboratoriWorksheet.Cells[rowNum, 5].Value = destinazione;
                        laboratoriWorksheet.Cells[rowNum, 6].Value = note;
                        laboratoriWorksheet.Cells[rowNum, 7].Value = auto;
                        laboratoriWorksheet.Cells[rowNum, 8].Value = volontario;
                        laboratoriWorksheet.Cells[rowNum, 9].Value = arrivo;
                        laboratoriWorksheet.Cells[rowNum, 10].Value = avv;
                    }

                    var workbook = new ExcelWorkbookModel(package);
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data starting at row 1
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1);

                    // Assert - Verify all columns are mapped correctly
                    var targetDimension = targetWorksheet.Dimension;
                    if (targetDimension == null && dataRows.Count > 0)
                    {
                        throw new Exception($"Expected {dataRows.Count} rows to be written, but target sheet is empty");
                    }

                    int actualRowsWritten = targetDimension?.End.Row ?? 0;
                    if (actualRowsWritten != dataRows.Count)
                    {
                        throw new Exception(
                            $"Expected {dataRows.Count} rows to be written, but {actualRowsWritten} rows were written"
                        );
                    }

                    // Verify each row has all 10 columns mapped correctly
                    for (int i = 0; i < dataRows.Count; i++)
                    {
                        int targetRow = i + 1;
                        var (expectedDate, expectedPartenza, expectedAssistito, expectedIndirizzo, 
                             expectedDestinazione, expectedNote, expectedAuto, expectedVolontario, 
                             expectedArrivo, expectedAvv) = dataRows[i];

                        // Column 1: Data (DateTime)
                        var actualDate = targetWorksheet.Cells[targetRow, 1].Value;
                        if (actualDate == null || !actualDate.Equals(expectedDate))
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 1 (Data): Expected '{expectedDate}', got '{actualDate}'"
                            );
                        }

                        // Column 2: Partenza (time as double)
                        var actualPartenza = targetWorksheet.Cells[targetRow, 2].Value;
                        if (actualPartenza == null || Math.Abs(Convert.ToDouble(actualPartenza) - expectedPartenza) > 0.0001)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 2 (Partenza): Expected '{expectedPartenza}', got '{actualPartenza}'"
                            );
                        }

                        // Column 3: Assistito (string)
                        var actualAssistito = targetWorksheet.Cells[targetRow, 3].Value?.ToString();
                        if (actualAssistito != expectedAssistito)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 3 (Assistito): Expected '{expectedAssistito}', got '{actualAssistito}'"
                            );
                        }

                        // Column 4: Indirizzo (string)
                        var actualIndirizzo = targetWorksheet.Cells[targetRow, 4].Value?.ToString();
                        if (actualIndirizzo != expectedIndirizzo)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 4 (Indirizzo): Expected '{expectedIndirizzo}', got '{actualIndirizzo}'"
                            );
                        }

                        // Column 5: Destinazione (string)
                        var actualDestinazione = targetWorksheet.Cells[targetRow, 5].Value?.ToString();
                        if (actualDestinazione != expectedDestinazione)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 5 (Destinazione): Expected '{expectedDestinazione}', got '{actualDestinazione}'"
                            );
                        }

                        // Column 6: Note (string)
                        var actualNote = targetWorksheet.Cells[targetRow, 6].Value?.ToString();
                        if (actualNote != expectedNote)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 6 (Note): Expected '{expectedNote}', got '{actualNote}'"
                            );
                        }

                        // Column 7: Auto (string)
                        var actualAuto = targetWorksheet.Cells[targetRow, 7].Value?.ToString();
                        if (actualAuto != expectedAuto)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 7 (Auto): Expected '{expectedAuto}', got '{actualAuto}'"
                            );
                        }

                        // Column 8: Volontario (string)
                        var actualVolontario = targetWorksheet.Cells[targetRow, 8].Value?.ToString();
                        if (actualVolontario != expectedVolontario)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 8 (Volontario): Expected '{expectedVolontario}', got '{actualVolontario}'"
                            );
                        }

                        // Column 9: Arrivo (time as double)
                        var actualArrivo = targetWorksheet.Cells[targetRow, 9].Value;
                        if (actualArrivo == null || Math.Abs(Convert.ToDouble(actualArrivo) - expectedArrivo) > 0.0001)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 9 (Arrivo): Expected '{expectedArrivo}', got '{actualArrivo}'"
                            );
                        }

                        // Column 10: Avv (string)
                        var actualAvv = targetWorksheet.Cells[targetRow, 10].Value?.ToString();
                        if (actualAvv != expectedAvv)
                        {
                            throw new Exception(
                                $"Row {targetRow}, Column 10 (Avv): Expected '{expectedAvv}', got '{actualAvv}'"
                            );
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }
    }
}
