using System;
using System.Collections.Generic;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Preservation property tests for laboratori data not appearing bugfix.
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
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// </summary>
    [TestFixture]
    public class LaboratoriDataPreservationTests
    {
        private IExcelManager _excelManager = null!;
        private IDataTransformer _dataTransformer = null!;
        private ICSVParser _csvParser = null!;
        private ILookupService _lookupService = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _excelManager = new ExcelManager();
            _dataTransformer = new DataTransformer(new TransformationRulesEngine());
            _csvParser = new CSVParser();
            _lookupService = new LookupService();
        }

        /// <summary>
        /// Property 2: Preservation - Files Without Laboratori Sheet Process Correctly
        /// 
        /// For any Excel file without a laboratori sheet, the processing workflow SHALL
        /// continue to complete successfully, processing only CSV and fissi data without
        /// errors or exceptions.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_FilesWithoutLaboratoriSheet()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of CSV rows (1-5)
            var csvRowCountGen = Gen.Choose(1, 5);

            // Generator for number of fissi rows (1-5)
            var fissiRowCountGen = Gen.Choose(1, 5);

            var testGen = from csvRowCount in csvRowCountGen
                          from fissiRowCount in fissiRowCountGen
                          select (csvRowCount, fissiRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (csvRowCount, fissiRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with multiple rows
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    
                    for (int i = 0; i < fissiRowCount; i++)
                    {
                        fissiWorksheet.Cells[2 + i, 1].Value = new DateTime(2026, 2, 1 + i);
                        fissiWorksheet.Cells[2 + i, 2].Value = 0.375 + (i * 0.01);
                        fissiWorksheet.Cells[2 + i, 3].Value = $"Fissi{i}";
                    }

                    // NO laboratori sheet created - this is the preservation condition

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Create CSV data
                    var csvAppointments = new List<ServiceAppointment>();
                    for (int i = 0; i < csvRowCount; i++)
                    {
                        csvAppointments.Add(new ServiceAppointment
                        {
                            DataServizio = $"{3 + i}/02/2026",
                            OraInizioServizio = "14:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        });
                    }

                    // Act - Transform and write data
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                    
                    int fissiStartRow = 3 + transformedResult.Rows.Count;
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                    
                    int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;

                    // Try to get laboratori sheet (should be null)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    
                    // Simulate ApplicationController workflow - only append if sheet exists
                    if (laboratoriSheetModel != null)
                    {
                        int laboratoriStartRow = lastRowAfterFissi + 1;
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                    }
                    
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Assert - Verify workflow completes without errors
                    if (laboratoriSheetModel != null)
                    {
                        throw new Exception("Laboratori sheet should not exist in this test");
                    }

                    // Verify only CSV + fissi data (no laboratori)
                    int expectedMinDataRows = csvRowCount + fissiRowCount;
                    int actualDataRows = finalLastRow - 2; // Subtract header rows
                    
                    // Should have at least the expected rows (may have more due to transformation)
                    if (actualDataRows < expectedMinDataRows)
                    {
                        throw new Exception(
                            $"Expected at least {expectedMinDataRows} data rows (CSV + fissi), " +
                            $"but found {actualDataRows}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Empty Laboratori Sheet Handling
        /// 
        /// For any Excel file with an empty laboratori sheet (null Dimension), the processing
        /// workflow SHALL continue to complete successfully without errors, processing only
        /// CSV and fissi data.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.2**
        /// </summary>
        [Test]
        public void Property_Preservation_EmptyLaboratoriSheet()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of CSV rows (1-5)
            var csvRowCountGen = Gen.Choose(1, 5);

            // Generator for number of fissi rows (1-5)
            var fissiRowCountGen = Gen.Choose(1, 5);

            var testGen = from csvRowCount in csvRowCountGen
                          from fissiRowCount in fissiRowCountGen
                          select (csvRowCount, fissiRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (csvRowCount, fissiRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with multiple rows
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    
                    for (int i = 0; i < fissiRowCount; i++)
                    {
                        fissiWorksheet.Cells[2 + i, 1].Value = new DateTime(2026, 2, 1 + i);
                        fissiWorksheet.Cells[2 + i, 2].Value = 0.375 + (i * 0.01);
                        fissiWorksheet.Cells[2 + i, 3].Value = $"Fissi{i}";
                    }

                    // Create EMPTY laboratori sheet (no data, null Dimension)
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    // Don't add any data - sheet exists but is empty

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Create CSV data
                    var csvAppointments = new List<ServiceAppointment>();
                    for (int i = 0; i < csvRowCount; i++)
                    {
                        csvAppointments.Add(new ServiceAppointment
                        {
                            DataServizio = $"{3 + i}/02/2026",
                            OraInizioServizio = "14:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        });
                    }

                    // Act - Transform and write data
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                    
                    int fissiStartRow = 3 + transformedResult.Rows.Count;
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                    
                    int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;

                    // Get laboratori sheet (should exist but be empty)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    
                    // Simulate ApplicationController workflow - only append if sheet exists
                    bool exceptionThrown = false;
                    if (laboratoriSheetModel != null)
                    {
                        try
                        {
                            int laboratoriStartRow = lastRowAfterFissi + 1;
                            _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                        }
                        catch (Exception)
                        {
                            exceptionThrown = true;
                        }
                    }
                    
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Assert - Verify workflow completes without errors
                    if (exceptionThrown)
                    {
                        throw new Exception("Empty laboratori sheet should not cause exceptions");
                    }

                    // Verify only CSV + fissi data (no laboratori)
                    int expectedMinDataRows = csvRowCount + fissiRowCount;
                    int actualDataRows = finalLastRow - 2; // Subtract header rows
                    
                    // Should have at least the expected rows
                    if (actualDataRows < expectedMinDataRows)
                    {
                        throw new Exception(
                            $"Expected at least {expectedMinDataRows} data rows (CSV + fissi), " +
                            $"but found {actualDataRows}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Fissi-Only Processing Produces Same Output
        /// 
        /// For any Excel file with only fissi data (no CSV data, no laboratori sheet),
        /// the processing workflow SHALL continue to produce the same output as before.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1, 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_FissiOnlyProcessing()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of fissi rows (1-10)
            var fissiRowCountGen = Gen.Choose(1, 10);

            Prop.ForAll(Arb.From(fissiRowCountGen), fissiRowCount =>
            {
                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with multiple rows
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    
                    for (int i = 0; i < fissiRowCount; i++)
                    {
                        fissiWorksheet.Cells[2 + i, 1].Value = new DateTime(2026, 2, 1 + i);
                        fissiWorksheet.Cells[2 + i, 2].Value = 0.375 + (i * 0.01);
                        fissiWorksheet.Cells[2 + i, 3].Value = $"Fissi{i}";
                    }

                    // NO laboratori sheet, NO CSV data - fissi only

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Act - Write headers and fissi data only (no CSV)
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, 3, DateTime.Now);
                    
                    int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? 2;

                    // Try to get laboratori sheet (should be null)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    
                    // Simulate ApplicationController workflow
                    if (laboratoriSheetModel != null)
                    {
                        int laboratoriStartRow = lastRowAfterFissi + 1;
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                    }
                    
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Assert - Verify only fissi data is present
                    int actualDataRows = finalLastRow - 2; // Subtract header rows
                    
                    if (actualDataRows != fissiRowCount)
                    {
                        throw new Exception(
                            $"Expected {fissiRowCount} fissi data rows, but found {actualDataRows}"
                        );
                    }

                    // Verify all rows contain fissi data markers
                    for (int row = 3; row <= finalLastRow; row++)
                    {
                        var assistitoValue = targetWorksheet.Cells[row, 3].Text;
                        if (!assistitoValue.StartsWith("Fissi"))
                        {
                            throw new Exception(
                                $"Row {row} should contain fissi data, but found: {assistitoValue}"
                            );
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - CSV-Only Processing Produces Same Output
        /// 
        /// For any CSV data without fissi or laboratori sheets, the processing workflow
        /// SHALL continue to produce the same output as before.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_CSVOnlyProcessing()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of CSV rows (1-10)
            var csvRowCountGen = Gen.Choose(1, 10);

            Prop.ForAll(Arb.From(csvRowCountGen), csvRowCount =>
            {
                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create minimal fissi sheet (for reference loading, but no data)
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    // No data rows - just headers

                    // NO laboratori sheet - CSV only

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Create CSV data
                    var csvAppointments = new List<ServiceAppointment>();
                    for (int i = 0; i < csvRowCount; i++)
                    {
                        csvAppointments.Add(new ServiceAppointment
                        {
                            DataServizio = $"{3 + i}/02/2026",
                            OraInizioServizio = $"{14 + i}:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        });
                    }

                    // Act - Transform and write CSV data only
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                    
                    int lastRowAfterCSV = targetWorksheet.Dimension?.End.Row ?? 2;

                    // Try to get laboratori sheet (should be null)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    
                    // Simulate ApplicationController workflow
                    if (laboratoriSheetModel != null)
                    {
                        int laboratoriStartRow = lastRowAfterCSV + 1;
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                    }
                    
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterCSV;

                    // Assert - Verify only CSV data is present
                    int actualDataRows = finalLastRow - 2; // Subtract header rows
                    
                    // Should have exactly the CSV rows (transformed)
                    if (actualDataRows != transformedResult.Rows.Count)
                    {
                        throw new Exception(
                            $"Expected {transformedResult.Rows.Count} CSV data rows, but found {actualDataRows}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Sorting By Date and Time Remains Unchanged
        /// 
        /// For any combination of CSV and fissi data (without laboratori), the sorting
        /// behavior SHALL continue to work correctly without throwing exceptions.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_SortingByDateAndTime()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of fissi rows (3-8)
            var fissiRowCountGen = Gen.Choose(3, 8);

            Prop.ForAll(Arb.From(fissiRowCountGen), fissiRowCount =>
            {
                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with data (will be sorted later)
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    
                    // Add rows with varying dates and times
                    for (int i = 0; i < fissiRowCount; i++)
                    {
                        fissiWorksheet.Cells[2 + i, 1].Value = new DateTime(2026, 2, 1 + (i % 5));
                        fissiWorksheet.Cells[2 + i, 2].Value = 0.25 + (i * 0.05); // Varying times
                        fissiWorksheet.Cells[2 + i, 3].Value = $"Fissi{i}";
                    }

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Act - Write headers and fissi data, then sort
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, 3, DateTime.Now);
                    
                    int lastRow = targetWorksheet.Dimension?.End.Row ?? 2;

                    // Try to get laboratori sheet (should be null)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    if (laboratoriSheetModel != null)
                    {
                        int laboratoriStartRow = lastRow + 1;
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                        lastRow = targetWorksheet.Dimension?.End.Row ?? lastRow;
                    }

                    // Sort the data - should not throw exceptions
                    bool exceptionThrown = false;
                    try
                    {
                        if (lastRow >= 3)
                        {
                            _excelManager.SortDataRows(targetSheet, 3, lastRow);
                        }
                    }
                    catch (Exception)
                    {
                        exceptionThrown = true;
                    }

                    // Assert - Sorting should work without exceptions
                    if (exceptionThrown)
                    {
                        throw new Exception("Sorting should work without exceptions when no laboratori data");
                    }

                    // Verify data is present after sorting
                    if (targetWorksheet.Dimension == null)
                    {
                        throw new Exception("Target sheet should have data after sorting");
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Time Format Handling for Partenza and Arrivo
        /// 
        /// For any laboratori sheet data with time values in Partenza (column 2) and Arrivo (column 9),
        /// the AppendLaboratoriData method SHALL continue to handle different time formats correctly
        /// (numeric, DateTime, string).
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_TimeFormatHandling()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for time format types
            var timeFormatGen = Gen.Elements("numeric", "datetime", "string");

            var testGen = from partenzaFormat in timeFormatGen
                          from arrivoFormat in timeFormatGen
                          select (partenzaFormat, arrivoFormat);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (partenzaFormat, arrivoFormat) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet with headers in row 1 (standard case)
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    laboratoriWorksheet.Cells[1, 1].Value = "Data";
                    laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                    laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                    laboratoriWorksheet.Cells[1, 9].Value = "Arrivo";

                    // Add data row with different time formats
                    laboratoriWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 1);
                    laboratoriWorksheet.Cells[2, 3].Value = "Test Patient";

                    // Set Partenza based on format type
                    switch (partenzaFormat)
                    {
                        case "numeric":
                            laboratoriWorksheet.Cells[2, 2].Value = 0.375; // 9:00 AM
                            break;
                        case "datetime":
                            laboratoriWorksheet.Cells[2, 2].Value = new DateTime(2026, 2, 1, 9, 0, 0);
                            break;
                        case "string":
                            laboratoriWorksheet.Cells[2, 2].Value = "09:00";
                            break;
                    }

                    // Set Arrivo based on format type
                    switch (arrivoFormat)
                    {
                        case "numeric":
                            laboratoriWorksheet.Cells[2, 9].Value = 0.5; // 12:00 PM
                            break;
                        case "datetime":
                            laboratoriWorksheet.Cells[2, 9].Value = new DateTime(2026, 2, 1, 12, 0, 0);
                            break;
                        case "string":
                            laboratoriWorksheet.Cells[2, 9].Value = "12:00";
                            break;
                    }

                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("target");
                    
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data
                    bool exceptionThrown = false;
                    try
                    {
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1, DateTime.Now);
                    }
                    catch (Exception)
                    {
                        exceptionThrown = true;
                    }

                    // Assert - Should not throw exceptions for any time format
                    if (exceptionThrown)
                    {
                        throw new Exception(
                            $"Time format handling should not throw exceptions. " +
                            $"Partenza format: {partenzaFormat}, Arrivo format: {arrivoFormat}"
                        );
                    }

                    // Verify data was copied
                    var targetDimension = targetWorksheet.Dimension;
                    if (targetDimension == null || targetDimension.End.Row < 1)
                    {
                        throw new Exception(
                            $"Expected data to be copied, but target sheet is empty. " +
                            $"Partenza format: {partenzaFormat}, Arrivo format: {arrivoFormat}"
                        );
                    }

                    // Verify time columns have values (format may vary)
                    var partenzaValue = targetWorksheet.Cells[1, 2].Value;
                    var arrivoValue = targetWorksheet.Cells[1, 9].Value;
                    
                    if (partenzaValue == null)
                    {
                        throw new Exception(
                            $"Partenza value should be copied. Format: {partenzaFormat}"
                        );
                    }
                    
                    if (arrivoValue == null)
                    {
                        throw new Exception(
                            $"Arrivo value should be copied. Format: {arrivoFormat}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Formatting Copy Excludes Background Colors
        /// 
        /// For any laboratori sheet data with cell formatting, the AppendLaboratoriData method
        /// SHALL continue to copy fonts, borders, and number formats while excluding background
        /// colors to avoid unintended yellow highlighting.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.4**
        /// </summary>
        [Test]
        public void Property_Preservation_FormattingCopyExcludesBackgroundColors()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 20;

            // Generator for background color presence
            var hasBackgroundColorGen = Arb.Generate<bool>();

            Prop.ForAll(Arb.From(hasBackgroundColorGen), hasBackgroundColor =>
            {
                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet with headers in row 1
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    laboratoriWorksheet.Cells[1, 1].Value = "Data";
                    laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                    laboratoriWorksheet.Cells[1, 3].Value = "Assistito";

                    // Add data row with formatting
                    var sourceCell = laboratoriWorksheet.Cells[2, 1];
                    sourceCell.Value = new DateTime(2026, 2, 1);
                    
                    // Apply various formatting
                    sourceCell.Style.Font.Bold = true;
                    sourceCell.Style.Font.Italic = true;
                    sourceCell.Style.Font.Size = 12;
                    sourceCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    sourceCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                    sourceCell.Style.Numberformat.Format = "dd/mm/yyyy";
                    
                    // Conditionally apply background color
                    if (hasBackgroundColor)
                    {
                        sourceCell.Style.Fill.PatternType = ExcelFillStyle.Solid;
                        sourceCell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    }

                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("target");
                    
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1, DateTime.Now);

                    // Assert - Verify formatting is copied correctly
                    var targetCell = targetWorksheet.Cells[1, 1];
                    
                    // Font properties should be copied
                    if (!targetCell.Style.Font.Bold)
                    {
                        throw new Exception("Font bold should be copied");
                    }
                    
                    if (!targetCell.Style.Font.Italic)
                    {
                        throw new Exception("Font italic should be copied");
                    }
                    
                    if (targetCell.Style.Font.Size != 12)
                    {
                        throw new Exception($"Font size should be 12, but got {targetCell.Style.Font.Size}");
                    }
                    
                    // Borders should be copied
                    if (targetCell.Style.Border.Top.Style != ExcelBorderStyle.Thin)
                    {
                        throw new Exception("Top border should be copied");
                    }
                    
                    // Number format should be copied
                    if (targetCell.Style.Numberformat.Format != "dd/mm/yyyy")
                    {
                        throw new Exception(
                            $"Number format should be 'dd/mm/yyyy', but got '{targetCell.Style.Numberformat.Format}'"
                        );
                    }
                    
                    // Background color should NOT be copied (requirement 3.4)
                    if (targetCell.Style.Fill.PatternType == ExcelFillStyle.Solid)
                    {
                        throw new Exception(
                            "Background color should NOT be copied to avoid yellow highlighting. " +
                            $"Source had background: {hasBackgroundColor}"
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Header Detection for Row 1 with "Data"
        /// 
        /// For any laboratori sheet where row 1 contains "Data" in cell A1 (standard header structure),
        /// the AppendLaboratoriData method SHALL continue to correctly identify row 2 as the data
        /// start row and process data accordingly.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1**
        /// </summary>
        [Test]
        public void Property_Preservation_HeaderDetectionRow1WithData()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            // Generator for number of data rows (1-10)
            var dataRowCountGen = Gen.Choose(1, 10);

            Prop.ForAll(Arb.From(dataRowCountGen), dataRowCount =>
            {
                using (var package = new ExcelPackage())
                {
                    // Create laboratori sheet with headers in row 1 (standard case)
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Row 1: Column headers with "Data" in A1
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

                    // Rows 2+: Data rows
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        int row = 2 + i;
                        laboratoriWorksheet.Cells[row, 1].Value = new DateTime(2026, 2, 1 + i);
                        laboratoriWorksheet.Cells[row, 2].Value = 0.333333 + (i * 0.01);
                        laboratoriWorksheet.Cells[row, 3].Value = $"Patient{i}";
                        laboratoriWorksheet.Cells[row, 4].Value = $"Address{i}";
                        laboratoriWorksheet.Cells[row, 5].Value = "Destination";
                        laboratoriWorksheet.Cells[row, 6].Value = $"Note{i}";
                        laboratoriWorksheet.Cells[row, 7].Value = $"Auto{i}";
                        laboratoriWorksheet.Cells[row, 8].Value = $"Vol{i}";
                        laboratoriWorksheet.Cells[row, 9].Value = 0.416667 + (i * 0.01);
                        laboratoriWorksheet.Cells[row, 10].Value = $"Avv{i}";
                    }

                    // Create target sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("target");
                    
                    var laboratoriSheet = new Sheet(laboratoriWorksheet);
                    var targetSheet = new Sheet(targetWorksheet);

                    // Act - Append laboratori data
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheet, 1, DateTime.Now);

                    // Assert - Verify all data rows were copied (header in row 1 should be skipped)
                    var targetDimension = targetWorksheet.Dimension;
                    if (targetDimension == null)
                    {
                        throw new Exception("Target sheet should have data after AppendLaboratoriData");
                    }

                    int actualRowsWritten = targetDimension.End.Row;
                    
                    if (actualRowsWritten != dataRowCount)
                    {
                        throw new Exception(
                            $"Expected {dataRowCount} data rows (header in row 1 should be skipped), " +
                            $"but found {actualRowsWritten} rows in target"
                        );
                    }

                    // Verify first row contains first patient (not header)
                    var firstPatient = targetWorksheet.Cells[1, 3].Value?.ToString();
                    if (firstPatient != "Patient0")
                    {
                        throw new Exception(
                            $"Expected first patient to be 'Patient0', but got '{firstPatient}'. " +
                            $"Header detection may have failed."
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property 2: Preservation - Complete Workflow Without Laboratori Data
        /// 
        /// For any complete processing workflow (CSV + fissi) without laboratori data,
        /// the output SHALL continue to be identical to the original behavior, including
        /// sorting, formatting, and all transformations.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_CompleteWorkflowWithoutLaboratori()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 20;

            // Generator for CSV row count (1-5)
            var csvRowCountGen = Gen.Choose(1, 5);

            // Generator for fissi row count (1-5)
            var fissiRowCountGen = Gen.Choose(1, 5);

            var testGen = from csvRowCount in csvRowCountGen
                          from fissiRowCount in fissiRowCountGen
                          select (csvRowCount, fissiRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (csvRowCount, fissiRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with multiple rows
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    
                    for (int i = 0; i < fissiRowCount; i++)
                    {
                        fissiWorksheet.Cells[2 + i, 1].Value = new DateTime(2026, 2, 1 + i);
                        fissiWorksheet.Cells[2 + i, 2].Value = 0.375 + (i * 0.01);
                        fissiWorksheet.Cells[2 + i, 3].Value = $"Fissi{i}";
                    }

                    // NO laboratori sheet

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Create CSV data
                    var csvAppointments = new List<ServiceAppointment>();
                    for (int i = 0; i < csvRowCount; i++)
                    {
                        csvAppointments.Add(new ServiceAppointment
                        {
                            DataServizio = $"{3 + i}/02/2026",
                            OraInizioServizio = $"{14 + i}:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        });
                    }

                    // Act - Complete workflow
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                    
                    int fissiStartRow = 3 + transformedResult.Rows.Count;
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                    
                    int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;

                    // Try to get laboratori sheet (should be null)
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    
                    // Simulate ApplicationController workflow
                    if (laboratoriSheetModel != null)
                    {
                        int laboratoriStartRow = lastRowAfterFissi + 1;
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                    }
                    
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Sort the data
                    if (finalLastRow >= 3)
                    {
                        _excelManager.SortDataRows(targetSheet, 3, finalLastRow);
                    }

                    // Assert - Verify complete workflow
                    if (laboratoriSheetModel != null)
                    {
                        throw new Exception("Laboratori sheet should not exist");
                    }

                    // Verify data is present
                    if (targetWorksheet.Dimension == null)
                    {
                        throw new Exception("Target sheet should have data");
                    }

                    // Verify sorting worked (dates should be in ascending order)
                    DateTime? previousDate = null;
                    for (int row = 3; row <= finalLastRow; row++)
                    {
                        var dateValue = targetWorksheet.Cells[row, 1].Value;
                        if (dateValue is DateTime currentDate)
                        {
                            if (previousDate.HasValue && currentDate < previousDate.Value)
                            {
                                throw new Exception(
                                    $"Row {row} date {currentDate:yyyy-MM-dd} is before previous date {previousDate.Value:yyyy-MM-dd}. " +
                                    $"Sorting should work correctly."
                                );
                            }
                            previousDate = currentDate;
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }
    }
}
