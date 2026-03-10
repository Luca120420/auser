using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using OfficeOpenXml;
using FsCheck;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Bug condition exploration test for laboratori data not appearing in output.
    /// **Validates: Requirements 2.1, 2.2, 2.4**
    /// 
    /// CRITICAL: This test is EXPECTED TO FAIL on unfixed code.
    /// Failure confirms the bug exists. Success after fix validates the correction.
    /// </summary>
    [TestFixture]
    public class LaboratoriBugExplorationTests
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
        /// Property 1: Fault Condition - Laboratori Data Appears in Output
        /// 
        /// For any Excel file where laboratori sheet exists with row 2 containing "Data" in A2 
        /// and rows 3+ containing data with non-empty Data values, the output file MUST contain 
        /// laboratoriRecordCount > 0.
        /// 
        /// **EXPECTED OUTCOME ON UNFIXED CODE**: Test FAILS with laboratoriRecordCount == 0
        /// This failure proves the bug exists.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.4**
        /// </summary>
        [Test]
        public void Property_LaboratoriDataAppearsInOutput_BugCondition()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            // Generator for number of laboratori data rows (1-10)
            var dataRowCountGen = Gen.Choose(1, 10);

            // Generator for number of empty rows to intersperse (0-3)
            var emptyRowCountGen = Gen.Choose(0, 3);

            var testGen = from dataRowCount in dataRowCountGen
                          from emptyRowCount in emptyRowCountGen
                          select (dataRowCount, emptyRowCount);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (dataRowCount, emptyRowCount) = tuple;

                using (var package = new ExcelPackage())
                {
                    Console.WriteLine($"\n=== DIAGNOSTIC LOG START ===");
                    Console.WriteLine($"Test parameters: dataRowCount={dataRowCount}, emptyRowCount={emptyRowCount}");

                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet (minimal)
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 2);
                    fissiWorksheet.Cells[2, 2].Value = 0.375;
                    fissiWorksheet.Cells[2, 3].Value = "Bianchi Laura";

                    // Create laboratori sheet with TWO-ROW HEADER structure (bug condition)
                    // Row 1: Metadata headers (e.g., "19 gen", "25 gen", "Settimana 4")
                    // Row 2: Column headers (starting with "Data" in A2)
                    // Row 3+: Actual data
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Row 1: Metadata headers (useless header)
                    laboratoriWorksheet.Cells[1, 1].Value = "19 gen";
                    laboratoriWorksheet.Cells[1, 2].Value = "25 gen";
                    laboratoriWorksheet.Cells[1, 3].Value = "Settimana 4";
                    
                    // Row 2: Column headers
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

                    // Row 3+: Data rows with some empty rows interspersed
                    int currentRow = 3;
                    int expectedNonEmptyRows = 0;
                    
                    for (int i = 0; i < dataRowCount; i++)
                    {
                        // Add a valid data row
                        laboratoriWorksheet.Cells[currentRow, 1].Value = new DateTime(2026, 2, 1 + i);
                        laboratoriWorksheet.Cells[currentRow, 2].Value = 0.333333 + (i * 0.01); // 8:00 AM + offset
                        laboratoriWorksheet.Cells[currentRow, 3].Value = $"Patient{i}";
                        laboratoriWorksheet.Cells[currentRow, 4].Value = $"Via Test {i}";
                        laboratoriWorksheet.Cells[currentRow, 5].Value = "Laboratorio";
                        laboratoriWorksheet.Cells[currentRow, 6].Value = $"Analisi {i}";
                        laboratoriWorksheet.Cells[currentRow, 7].Value = $"Auto{i}";
                        laboratoriWorksheet.Cells[currentRow, 8].Value = $"Vol{i}";
                        laboratoriWorksheet.Cells[currentRow, 9].Value = 0.416667 + (i * 0.01); // 10:00 AM + offset
                        laboratoriWorksheet.Cells[currentRow, 10].Value = $"Avv{i}";
                        expectedNonEmptyRows++;
                        currentRow++;

                        // Occasionally add an empty row (should be skipped)
                        if (i < emptyRowCount)
                        {
                            // Leave column 1 empty - this row should be skipped
                            laboratoriWorksheet.Cells[currentRow, 3].Value = "ShouldBeSkipped";
                            currentRow++;
                        }
                    }

                    Console.WriteLine($"Created laboratori sheet with {expectedNonEmptyRows} non-empty data rows");
                    Console.WriteLine($"Row 2, Col 1 (A2) contains: '{laboratoriWorksheet.Cells[2, 1].Text}'");

                    // Create target output sheet
                    var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                    
                    // Wrap sheets in models
                    var workbook = new Models.ExcelWorkbook(package);
                    
                    // DIAGNOSTIC: Log all worksheet names in workbook before GetSheetByName
                    Console.WriteLine("\n--- DIAGNOSTIC: Worksheet names in workbook ---");
                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        Console.WriteLine($"  - '{ws.Name}'");
                    }
                    
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    
                    // DIAGNOSTIC: Log result of GetSheetByName for "laboratori"
                    Console.WriteLine("\n--- DIAGNOSTIC: GetSheetByName for 'laboratori' ---");
                    var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    Console.WriteLine($"  Result: {(laboratoriSheetModel != null ? "NOT NULL (sheet found)" : "NULL (sheet NOT found)")}");
                    
                    var targetSheet = new Sheet(targetWorksheet);

                    // Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Create minimal CSV data
                    var csvAppointments = new List<ServiceAppointment>
                    {
                        new ServiceAppointment
                        {
                            DataServizio = "03/02/2026",
                            OraInizioServizio = "14:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            IndirizzoPartenza = "Via Roma 1",
                            IndirizzoDestinazione = "Via Milano 5",
                            ComunePartenza = "Milano",
                            ComuneDestinazione = "Milano",
                            DescrizionePuntoPartenza = "Casa",
                            CausaleDestinazione = "Ospedale",
                            NoteERichieste = "Test note",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        }
                    };

                    // Act - Transform CSV data
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    
                    // Write headers and CSV data
                    _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                    _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                    
                    // Append fissi data
                    int fissiStartRow = 3 + transformedResult.Rows.Count;
                    _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                    
                    // Calculate last row after fissi
                    int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                    
                    // DIAGNOSTIC: Log Dimension before AppendLaboratoriData
                    Console.WriteLine("\n--- DIAGNOSTIC: Before AppendLaboratoriData ---");
                    Console.WriteLine($"  targetWorksheet.Dimension: {(targetWorksheet.Dimension != null ? targetWorksheet.Dimension.Address : "NULL")}");
                    Console.WriteLine($"  lastRowAfterFissi: {lastRowAfterFissi}");
                    Console.WriteLine($"  laboratoriStartRow: {lastRowAfterFissi + 1}");
                    
                    // Append laboratori data (if sheet exists)
                    if (laboratoriSheetModel != null)
                    {
                        Console.WriteLine("\n--- DIAGNOSTIC: Calling AppendLaboratoriData ---");
                        int laboratoriStartRow = lastRowAfterFissi + 1;
                        
                        // Add diagnostic wrapper to trace entry/exit
                        int rowCountBefore = targetWorksheet.Dimension?.End.Row ?? 0;
                        Console.WriteLine($"  Entry: targetRow={laboratoriStartRow}, sourceRows={expectedNonEmptyRows}");
                        
                        _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                        
                        int rowCountAfter = targetWorksheet.Dimension?.End.Row ?? 0;
                        Console.WriteLine($"  Exit: rowCountBefore={rowCountBefore}, rowCountAfter={rowCountAfter}");
                        Console.WriteLine($"  Rows added: {rowCountAfter - rowCountBefore}");
                    }
                    else
                    {
                        Console.WriteLine("\n--- DIAGNOSTIC: laboratoriSheetModel is NULL, skipping AppendLaboratoriData ---");
                    }
                    
                    // DIAGNOSTIC: Log Dimension after AppendLaboratoriData
                    Console.WriteLine("\n--- DIAGNOSTIC: After AppendLaboratoriData ---");
                    Console.WriteLine($"  targetWorksheet.Dimension: {(targetWorksheet.Dimension != null ? targetWorksheet.Dimension.Address : "NULL")}");
                    
                    // Get final row count
                    int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Count laboratori records in output (rows with Avv column data)
                    int laboratoriRecordCount = 0;
                    for (int row = 3; row <= finalLastRow; row++)
                    {
                        var avvValue = targetWorksheet.Cells[row, 10].Text?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(avvValue) && avvValue.StartsWith("Avv"))
                        {
                            laboratoriRecordCount++;
                        }
                    }

                    Console.WriteLine($"\n--- DIAGNOSTIC: Final Results ---");
                    Console.WriteLine($"  Expected laboratori rows: {expectedNonEmptyRows}");
                    Console.WriteLine($"  Actual laboratori rows in output: {laboratoriRecordCount}");
                    Console.WriteLine($"  Total output rows: {finalLastRow}");
                    Console.WriteLine($"=== DIAGNOSTIC LOG END ===\n");

                    // Assert - This SHOULD FAIL on unfixed code (laboratoriRecordCount == 0)
                    // When it fails, it proves the bug exists
                    if (laboratoriRecordCount == 0)
                    {
                        throw new Exception(
                            $"BUG CONFIRMED: Expected {expectedNonEmptyRows} laboratori records in output, but found 0. " +
                            $"This confirms the bug exists. Check diagnostic logs above for root cause."
                        );
                    }

                    // Verify all expected rows are present
                    if (laboratoriRecordCount != expectedNonEmptyRows)
                    {
                        throw new Exception(
                            $"Expected {expectedNonEmptyRows} laboratori records in output, but found {laboratoriRecordCount}. " +
                            $"Some data rows are missing."
                        );
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Simplified single-case test that simulates the REAL ApplicationController workflow.
        /// Creates a real Excel file on disk, loads it, and processes it through ApplicationController.
        /// Tests with exactly 7 data rows as mentioned in the bug report.
        /// 
        /// **EXPECTED OUTCOME ON UNFIXED CODE**: Test FAILS with laboratoriRecordCount == 0
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.4**
        /// </summary>
        [Test]
        public void SimplifiedTest_SevenLaboratoriRows_BugCondition()
        {
            string testFilePath = Path.Combine(Path.GetTempPath(), $"test_laboratori_{Guid.NewGuid()}.xlsx");
            
            try
            {
                // Step 1: Create a real Excel file with laboratori sheet
                using (var package = new ExcelPackage())
                {
                    Console.WriteLine($"\n=== CREATING TEST FILE ===");
                    Console.WriteLine($"File path: {testFilePath}");

                    // Create assistiti sheet (reference data)
                    var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                    assistitiSheet.Cells[1, 1].Value = "Cognome";
                    assistitiSheet.Cells[1, 2].Value = "Nome";
                    assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                    assistitiSheet.Cells[2, 1].Value = "Rossi";
                    assistitiSheet.Cells[2, 2].Value = "Mario";
                    assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                    // Create fissi sheet with numbered sheet "1" for header reading
                    var sheet1 = package.Workbook.Worksheets.Add("1");
                    sheet1.Cells[1, 1].Value = "26 gen 01 feb Settimana 4";
                    sheet1.Cells[2, 1].Value = "Data";
                    
                    // Create fissi sheet (minimal)
                    var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                    fissiWorksheet.Cells[1, 1].Value = "Data";
                    fissiWorksheet.Cells[1, 2].Value = "Partenza";
                    fissiWorksheet.Cells[1, 3].Value = "Assistito";
                    fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 2);
                    fissiWorksheet.Cells[2, 2].Value = 0.375;
                    fissiWorksheet.Cells[2, 3].Value = "Bianchi Laura";

                    // Create laboratori sheet with TWO-ROW HEADER structure (bug condition)
                    // Row 1: Metadata headers
                    // Row 2: Column headers with "Data" in A2
                    // Rows 3-9: 7 data rows with non-empty Data values
                    // Rows 10-14: 5 rows with empty Data values (should be skipped)
                    var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                    
                    // Row 1: Metadata headers
                    laboratoriWorksheet.Cells[1, 1].Value = "19 gen";
                    laboratoriWorksheet.Cells[1, 2].Value = "25 gen";
                    laboratoriWorksheet.Cells[1, 3].Value = "Settimana 4";
                    
                    // Row 2: Column headers
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

                    // Rows 3-9: 7 data rows with non-empty Data values
                    for (int i = 0; i < 7; i++)
                    {
                        int row = 3 + i;
                        laboratoriWorksheet.Cells[row, 1].Value = new DateTime(2026, 1, 19 + i);
                        laboratoriWorksheet.Cells[row, 2].Value = 0.333333 + (i * 0.01);
                        laboratoriWorksheet.Cells[row, 3].Value = $"Ferrari Anna {i}";
                        laboratoriWorksheet.Cells[row, 4].Value = $"Via Napoli {20 + i}";
                        laboratoriWorksheet.Cells[row, 5].Value = "Laboratorio Analisi";
                        laboratoriWorksheet.Cells[row, 6].Value = $"Prelievo sangue {i}";
                        laboratoriWorksheet.Cells[row, 7].Value = $"Auto{i + 1}";
                        laboratoriWorksheet.Cells[row, 8].Value = $"Volontario{i + 1}";
                        laboratoriWorksheet.Cells[row, 9].Value = 0.416667 + (i * 0.01);
                        laboratoriWorksheet.Cells[row, 10].Value = $"Avviso{i + 1}";
                    }

                    // Rows 10-14: 5 rows with empty Data values (should be skipped)
                    for (int i = 0; i < 5; i++)
                    {
                        int row = 10 + i;
                        // Leave column 1 empty
                        laboratoriWorksheet.Cells[row, 3].Value = "EmptyDataRow";
                    }

                    Console.WriteLine($"Created laboratori sheet:");
                    Console.WriteLine($"  Row 1: Metadata headers");
                    Console.WriteLine($"  Row 2: Column headers with 'Data' in A2");
                    Console.WriteLine($"  Rows 3-9: 7 data rows with non-empty Data values");
                    Console.WriteLine($"  Rows 10-14: 5 rows with empty Data values (should be skipped)");

                    // Save the file
                    package.SaveAs(new FileInfo(testFilePath));
                    Console.WriteLine($"✓ Test file saved");
                }

                // Step 2: Load the file and process it through ApplicationController workflow
                Console.WriteLine($"\n=== PROCESSING THROUGH APPLICATIONCONTROLLER WORKFLOW ===");
                
                using (var package = new ExcelPackage(new FileInfo(testFilePath)))
                {
                    var workbook = new Models.ExcelWorkbook(package);
                    
                    // DIAGNOSTIC: Log all worksheet names in workbook after loading
                    Console.WriteLine("\n--- DIAGNOSTIC: Worksheet names in loaded workbook ---");
                    foreach (var ws in package.Workbook.Worksheets)
                    {
                        Console.WriteLine($"  - '{ws.Name}'");
                    }
                    
                    // Simulate ApplicationController workflow
                    // Step 3: Get reference sheets
                    var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                    var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                    
                    // Step 4: Load reference sheets
                    _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                    // Step 5: Create CSV data
                    var csvAppointments = new List<ServiceAppointment>
                    {
                        new ServiceAppointment
                        {
                            DataServizio = "03/02/2026",
                            OraInizioServizio = "14:00",
                            CognomeAssistito = "Rossi",
                            NomeAssistito = "Mario",
                            IndirizzoPartenza = "Via Roma 1",
                            IndirizzoDestinazione = "Via Milano 5",
                            ComunePartenza = "Milano",
                            ComuneDestinazione = "Milano",
                            DescrizionePuntoPartenza = "Casa",
                            CausaleDestinazione = "Ospedale",
                            NoteERichieste = "Test note",
                            DescrizioneStatoServizio = "PIANIFICATO",
                            Attivita = "Accomp. servizi con trasporto"
                        }
                    };

                    // Step 6: Transform CSV data
                    var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                    
                    // Step 7: Get next sheet number
                    var sheetNames = _excelManager.GetSheetNames(workbook);
                    int nextSheetNumber = _excelManager.GetNextSheetNumber(sheetNames);
                    
                    // Step 8: Create new sheet (THIS ADDS TO THE SAME WORKBOOK)
                    Sheet newSheet = _excelManager.CreateNewSheet(workbook, nextSheetNumber);
                    
                    Console.WriteLine($"\n--- DIAGNOSTIC: After CreateNewSheet ---");
                    Console.WriteLine($"  New sheet number: {nextSheetNumber}");
                    Console.WriteLine($"  Total worksheets in workbook: {package.Workbook.Worksheets.Count}");
                    
                    // Step 9: Write headers and CSV data
                    _excelManager.WriteHeaderRow(newSheet, new DateTime(2026, 2, 3));
                    _excelManager.WriteColumnHeadersEnhanced(newSheet);
                    _excelManager.ApplyBoldToHeaders(newSheet, 2);
                    _excelManager.WriteDataRowsEnhanced(newSheet, transformedResult.Rows, 3);
                    
                    // Step 10: Append fissi data
                    int fissiStartRow = 3 + transformedResult.Rows.Count;
                    _excelManager.AppendFissiData(newSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                    
                    // Calculate last row after fissi
                    int lastRowAfterFissi = newSheet.Worksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                    
                    // DIAGNOSTIC: Log Dimension before AppendLaboratoriData
                    Console.WriteLine("\n--- DIAGNOSTIC: Before AppendLaboratoriData ---");
                    Console.WriteLine($"  newSheet.Worksheet.Dimension: {(newSheet.Worksheet.Dimension != null ? newSheet.Worksheet.Dimension.Address : "NULL")}");
                    Console.WriteLine($"  lastRowAfterFissi: {lastRowAfterFissi}");
                    Console.WriteLine($"  laboratoriStartRow: {lastRowAfterFissi + 1}");
                    
                    // Step 10.25: Append laboratori data (EXACT CODE FROM ApplicationController)
                    Console.WriteLine("\n--- DIAGNOSTIC: GetSheetByName for 'laboratori' ---");
                    Sheet laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                    Console.WriteLine($"  Result: {(laboratoriSheetModel != null ? "NOT NULL (sheet found)" : "NULL (sheet NOT found)")}");
                    
                    if (laboratoriSheetModel != null)
                    {
                        Console.WriteLine("\n--- DIAGNOSTIC: Calling AppendLaboratoriData ---");
                        int laboratoriStartRow = lastRowAfterFissi + 1;
                        
                        int rowCountBefore = newSheet.Worksheet.Dimension?.End.Row ?? 0;
                        Console.WriteLine($"  Entry: targetRow={laboratoriStartRow}, expectedSourceRows=7");
                        
                        _excelManager.AppendLaboratoriData(newSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                        
                        int rowCountAfter = newSheet.Worksheet.Dimension?.End.Row ?? 0;
                        Console.WriteLine($"  Exit: rowCountBefore={rowCountBefore}, rowCountAfter={rowCountAfter}");
                        Console.WriteLine($"  Rows added: {rowCountAfter - rowCountBefore}");
                        
                        // Update last data row
                        lastRowAfterFissi = newSheet.Worksheet.Dimension?.End.Row ?? lastRowAfterFissi;
                    }
                    else
                    {
                        Console.WriteLine("\n--- DIAGNOSTIC: laboratoriSheetModel is NULL, skipping AppendLaboratoriData ---");
                    }
                    
                    // DIAGNOSTIC: Log Dimension after AppendLaboratoriData
                    Console.WriteLine("\n--- DIAGNOSTIC: After AppendLaboratoriData ---");
                    Console.WriteLine($"  newSheet.Worksheet.Dimension: {(newSheet.Worksheet.Dimension != null ? newSheet.Worksheet.Dimension.Address : "NULL")}");
                    
                    // Get final row count
                    int finalLastRow = newSheet.Worksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                    // Count laboratori records in output (rows with Avv column data)
                    int laboratoriRecordCount = 0;
                    for (int row = 3; row <= finalLastRow; row++)
                    {
                        var avvValue = newSheet.Worksheet.Cells[row, 10].Text?.Trim() ?? "";
                        if (!string.IsNullOrWhiteSpace(avvValue) && avvValue.StartsWith("Avv"))
                        {
                            laboratoriRecordCount++;
                        }
                    }

                    Console.WriteLine($"\n--- DIAGNOSTIC: Final Results ---");
                    Console.WriteLine($"  Expected laboratori rows: 7");
                    Console.WriteLine($"  Actual laboratori rows in output: {laboratoriRecordCount}");
                    Console.WriteLine($"  Total output rows: {finalLastRow}");
                    Console.WriteLine($"=== DIAGNOSTIC LOG END ===\n");

                    // Assert - This SHOULD FAIL on unfixed code (laboratoriRecordCount == 0)
                    // When it fails, it proves the bug exists
                    Assert.That(laboratoriRecordCount, Is.GreaterThan(0),
                        $"BUG CONFIRMED: Expected 7 laboratori records in output, but found {laboratoriRecordCount}. " +
                        $"This confirms the bug exists. Check diagnostic logs above for root cause.");

                    // Verify all expected rows are present
                    Assert.That(laboratoriRecordCount, Is.EqualTo(7),
                        $"Expected 7 laboratori records in output, but found {laboratoriRecordCount}. " +
                        $"Some data rows are missing.");
                }
            }
            finally
            {
                // Cleanup
                if (File.Exists(testFilePath))
                {
                    try
                    {
                        File.Delete(testFilePath);
                        Console.WriteLine($"✓ Test file deleted");
                    }
                    catch
                    {
                        // Ignore cleanup errors
                    }
                }
            }
        }

        /// <summary>
        /// Original in-memory test (kept for comparison).
        /// This test PASSES because all sheets are in the same in-memory workbook.
        /// </summary>
        [Test]
        public void InMemoryTest_SevenLaboratoriRows_WorksCorrectly()
        {
            using (var package = new ExcelPackage())
            {
                Console.WriteLine($"\n=== DIAGNOSTIC LOG START (Simplified Test) ===");

                // Create assistiti sheet (reference data)
                var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                assistitiSheet.Cells[1, 1].Value = "Cognome";
                assistitiSheet.Cells[1, 2].Value = "Nome";
                assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                assistitiSheet.Cells[2, 1].Value = "Rossi";
                assistitiSheet.Cells[2, 2].Value = "Mario";
                assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                // Create fissi sheet (minimal)
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                fissiWorksheet.Cells[1, 1].Value = "Data";
                fissiWorksheet.Cells[1, 2].Value = "Partenza";
                fissiWorksheet.Cells[1, 3].Value = "Assistito";
                fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 2);
                fissiWorksheet.Cells[2, 2].Value = 0.375;
                fissiWorksheet.Cells[2, 3].Value = "Bianchi Laura";

                // Create laboratori sheet with TWO-ROW HEADER structure (bug condition)
                // Row 1: Metadata headers
                // Row 2: Column headers with "Data" in A2
                // Rows 3-9: 7 data rows with non-empty Data values
                // Rows 10-14: 5 rows with empty Data values (should be skipped)
                var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                
                // Row 1: Metadata headers
                laboratoriWorksheet.Cells[1, 1].Value = "19 gen";
                laboratoriWorksheet.Cells[1, 2].Value = "25 gen";
                laboratoriWorksheet.Cells[1, 3].Value = "Settimana 4";
                
                // Row 2: Column headers
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

                // Rows 3-9: 7 data rows with non-empty Data values
                for (int i = 0; i < 7; i++)
                {
                    int row = 3 + i;
                    laboratoriWorksheet.Cells[row, 1].Value = new DateTime(2026, 1, 19 + i);
                    laboratoriWorksheet.Cells[row, 2].Value = 0.333333 + (i * 0.01);
                    laboratoriWorksheet.Cells[row, 3].Value = $"Ferrari Anna {i}";
                    laboratoriWorksheet.Cells[row, 4].Value = $"Via Napoli {20 + i}";
                    laboratoriWorksheet.Cells[row, 5].Value = "Laboratorio Analisi";
                    laboratoriWorksheet.Cells[row, 6].Value = $"Prelievo sangue {i}";
                    laboratoriWorksheet.Cells[row, 7].Value = $"Auto{i + 1}";
                    laboratoriWorksheet.Cells[row, 8].Value = $"Volontario{i + 1}";
                    laboratoriWorksheet.Cells[row, 9].Value = 0.416667 + (i * 0.01);
                    laboratoriWorksheet.Cells[row, 10].Value = $"Avviso{i + 1}";
                }

                // Rows 10-14: 5 rows with empty Data values (should be skipped)
                for (int i = 0; i < 5; i++)
                {
                    int row = 10 + i;
                    // Leave column 1 empty
                    laboratoriWorksheet.Cells[row, 3].Value = "EmptyDataRow";
                }

                Console.WriteLine($"Created laboratori sheet:");
                Console.WriteLine($"  Row 1: Metadata headers");
                Console.WriteLine($"  Row 2: Column headers with 'Data' in A2");
                Console.WriteLine($"  Rows 3-9: 7 data rows with non-empty Data values");
                Console.WriteLine($"  Rows 10-14: 5 rows with empty Data values (should be skipped)");

                // Create target output sheet
                var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                
                // Wrap sheets in models
                var workbook = new Models.ExcelWorkbook(package);
                
                // DIAGNOSTIC: Log all worksheet names in workbook before GetSheetByName
                Console.WriteLine("\n--- DIAGNOSTIC: Worksheet names in workbook ---");
                foreach (var ws in package.Workbook.Worksheets)
                {
                    Console.WriteLine($"  - '{ws.Name}'");
                }
                
                var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                
                // DIAGNOSTIC: Log result of GetSheetByName for "laboratori"
                Console.WriteLine("\n--- DIAGNOSTIC: GetSheetByName for 'laboratori' ---");
                var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                Console.WriteLine($"  Result: {(laboratoriSheetModel != null ? "NOT NULL (sheet found)" : "NULL (sheet NOT found)")}");
                
                var targetSheet = new Sheet(targetWorksheet);

                // Load reference sheets
                _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                // Create minimal CSV data
                var csvAppointments = new List<ServiceAppointment>
                {
                    new ServiceAppointment
                    {
                        DataServizio = "03/02/2026",
                        OraInizioServizio = "14:00",
                        CognomeAssistito = "Rossi",
                        NomeAssistito = "Mario",
                        IndirizzoPartenza = "Via Roma 1",
                        IndirizzoDestinazione = "Via Milano 5",
                        ComunePartenza = "Milano",
                        ComuneDestinazione = "Milano",
                        DescrizionePuntoPartenza = "Casa",
                        CausaleDestinazione = "Ospedale",
                        NoteERichieste = "Test note",
                        DescrizioneStatoServizio = "PIANIFICATO",
                        Attivita = "Accomp. servizi con trasporto"
                    }
                };

                // Act - Transform CSV data
                var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                
                // Write headers and CSV data
                _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                
                // Append fissi data
                int fissiStartRow = 3 + transformedResult.Rows.Count;
                _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow, DateTime.Now);
                
                // Calculate last row after fissi
                int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                
                // DIAGNOSTIC: Log Dimension before AppendLaboratoriData
                Console.WriteLine("\n--- DIAGNOSTIC: Before AppendLaboratoriData ---");
                Console.WriteLine($"  targetWorksheet.Dimension: {(targetWorksheet.Dimension != null ? targetWorksheet.Dimension.Address : "NULL")}");
                Console.WriteLine($"  lastRowAfterFissi: {lastRowAfterFissi}");
                Console.WriteLine($"  laboratoriStartRow: {lastRowAfterFissi + 1}");
                
                // Append laboratori data (if sheet exists)
                if (laboratoriSheetModel != null)
                {
                    Console.WriteLine("\n--- DIAGNOSTIC: Calling AppendLaboratoriData ---");
                    int laboratoriStartRow = lastRowAfterFissi + 1;
                    
                    // Add diagnostic wrapper to trace entry/exit
                    int rowCountBefore = targetWorksheet.Dimension?.End.Row ?? 0;
                    Console.WriteLine($"  Entry: targetRow={laboratoriStartRow}, expectedSourceRows=7");
                    
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow, DateTime.Now);
                    
                    int rowCountAfter = targetWorksheet.Dimension?.End.Row ?? 0;
                    Console.WriteLine($"  Exit: rowCountBefore={rowCountBefore}, rowCountAfter={rowCountAfter}");
                    Console.WriteLine($"  Rows added: {rowCountAfter - rowCountBefore}");
                }
                else
                {
                    Console.WriteLine("\n--- DIAGNOSTIC: laboratoriSheetModel is NULL, skipping AppendLaboratoriData ---");
                }
                
                // DIAGNOSTIC: Log Dimension after AppendLaboratoriData
                Console.WriteLine("\n--- DIAGNOSTIC: After AppendLaboratoriData ---");
                Console.WriteLine($"  targetWorksheet.Dimension: {(targetWorksheet.Dimension != null ? targetWorksheet.Dimension.Address : "NULL")}");
                
                // Get final row count
                int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;

                // Count laboratori records in output (rows with Avv column data)
                int laboratoriRecordCount = 0;
                for (int row = 3; row <= finalLastRow; row++)
                {
                    var avvValue = targetWorksheet.Cells[row, 10].Text?.Trim() ?? "";
                    if (!string.IsNullOrWhiteSpace(avvValue) && avvValue.StartsWith("Avv"))
                    {
                        laboratoriRecordCount++;
                    }
                }

                Console.WriteLine($"\n--- DIAGNOSTIC: Final Results ---");
                Console.WriteLine($"  Expected laboratori rows: 7");
                Console.WriteLine($"  Actual laboratori rows in output: {laboratoriRecordCount}");
                Console.WriteLine($"  Total output rows: {finalLastRow}");
                Console.WriteLine($"=== DIAGNOSTIC LOG END ===\n");

                // Assert - This SHOULD FAIL on unfixed code (laboratoriRecordCount == 0)
                // When it fails, it proves the bug exists
                Assert.That(laboratoriRecordCount, Is.GreaterThan(0),
                    $"BUG CONFIRMED: Expected 7 laboratori records in output, but found {laboratoriRecordCount}. " +
                    $"This confirms the bug exists. Check diagnostic logs above for root cause.");

                // Verify all expected rows are present
                Assert.That(laboratoriRecordCount, Is.EqualTo(7),
                    $"Expected 7 laboratori records in output, but found {laboratoriRecordCount}. " +
                    $"Some data rows are missing.");
            }
        }
    }
}
