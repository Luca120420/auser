using System;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Bug condition exploration tests for fissi column mapping fix.
    /// 
    /// CRITICAL: These tests are EXPECTED TO FAIL on unfixed code.
    /// Test failure confirms the bug exists.
    /// 
    /// These tests encode the EXPECTED (correct) behavior.
    /// When the bug is fixed, these tests will pass.
    /// 
    /// **Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.10, 2.11, 2.12, 2.13, 2.14, 2.15**
    /// </summary>
    [TestFixture]
    public class FissiColumnMappingBugExplorationTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Property 1: Fault Condition - Incorrect Column Mapping for Fissi Columns 4-9
        /// 
        /// This test creates a fissi sheet with known data in columns 1-9 and verifies
        /// that the data appears in the CORRECT target columns.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL because:
        /// - Fissi column 5 "Destinazione" currently goes to target column 4 (should be 5)
        /// - Fissi column 6 "Note" currently goes to target column 5 (should be 6)
        /// - Fissi column 7 "Auto" currently goes to target column 6 (should be 7)
        /// - Fissi column 8 "Volontario" currently goes to target column 7 (should be 8)
        /// - Fissi column 9 "Arrivo" currently goes to target column 8 (should be 9)
        /// - Fissi column 4 "Indirizzo" currently goes to target column 10 (should be skipped)
        /// - Target columns 10-15 should be empty but column 10 contains fissi column 4 data
        /// 
        /// When the bug is fixed, this test will PASS.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.10, 2.11, 2.12, 2.13, 2.14, 2.15**
        /// </summary>
        [Test]
        public void BugExploration_FissiColumnMapping_ShouldMapColumnsCorrectly()
        {
            // Arrange - Create a fissi sheet with known data in columns 1-9
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row in fissi sheet (row 2 contains "Data" as per header detection logic)
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";
                fissiWorksheet.Cells[2, 3].Value = "Assistito";
                fissiWorksheet.Cells[2, 4].Value = "Indirizzo";
                fissiWorksheet.Cells[2, 5].Value = "Destinazione";
                fissiWorksheet.Cells[2, 6].Value = "Note";
                fissiWorksheet.Cells[2, 7].Value = "Auto";
                fissiWorksheet.Cells[2, 8].Value = "Volontario";
                fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                // Add data row with known values in fissi sheet (row 3)
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";        // Data
                fissiWorksheet.Cells[3, 2].Value = "09:00";             // Partenza
                fissiWorksheet.Cells[3, 3].Value = "Rossi Mario";       // Assistito
                fissiWorksheet.Cells[3, 4].Value = "Via Roma 10";       // Indirizzo (should be SKIPPED)
                fissiWorksheet.Cells[3, 5].Value = "Ospedale";          // Destinazione
                fissiWorksheet.Cells[3, 6].Value = "Urgente";           // Note
                fissiWorksheet.Cells[3, 7].Value = "Auto1";             // Auto
                fissiWorksheet.Cells[3, 8].Value = "Mario Rossi";       // Volontario
                fissiWorksheet.Cells[3, 9].Value = "10:00";             // Arrivo

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act - Call AppendFissiData on UNFIXED code
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Verify CORRECT mapping (will fail on unfixed code)
                
                // Requirement 2.1: Fissi column 5 "Destinazione" should appear in target column 5
                var destinazioneValue = targetWorksheet.Cells[1, 5].Value?.ToString();
                Assert.That(destinazioneValue, Is.EqualTo("Ospedale"),
                    "COUNTEREXAMPLE: Fissi column 5 'Destinazione' should map to target column 5. " +
                    $"Expected 'Ospedale' in target column 5, but found '{destinazioneValue}'. " +
                    "This confirms the bug: Destinazione is being placed in the wrong column.");

                // Requirement 2.2: Fissi column 6 "Note" should appear in target column 6
                var noteValue = targetWorksheet.Cells[1, 6].Value?.ToString();
                Assert.That(noteValue, Is.EqualTo("Urgente"),
                    "COUNTEREXAMPLE: Fissi column 6 'Note' should map to target column 6. " +
                    $"Expected 'Urgente' in target column 6, but found '{noteValue}'. " +
                    "This confirms the bug: Note is being placed in the wrong column.");

                // Requirement 2.3: Fissi column 7 "Auto" should appear in target column 7
                var autoValue = targetWorksheet.Cells[1, 7].Value?.ToString();
                Assert.That(autoValue, Is.EqualTo("Auto1"),
                    "COUNTEREXAMPLE: Fissi column 7 'Auto' should map to target column 7. " +
                    $"Expected 'Auto1' in target column 7, but found '{autoValue}'. " +
                    "This confirms the bug: Auto is being placed in the wrong column.");

                // Requirement 2.4: Fissi column 8 "Volontario" should appear in target column 8
                var volontarioValue = targetWorksheet.Cells[1, 8].Value?.ToString();
                Assert.That(volontarioValue, Is.EqualTo("Mario Rossi"),
                    "COUNTEREXAMPLE: Fissi column 8 'Volontario' should map to target column 8. " +
                    $"Expected 'Mario Rossi' in target column 8, but found '{volontarioValue}'. " +
                    "This confirms the bug: Volontario is being placed in the wrong column.");

                // Requirement 2.5: Fissi column 9 "Arrivo" should appear in target column 9
                var arrivoValue = targetWorksheet.Cells[1, 9].Value?.ToString();
                Assert.That(arrivoValue, Is.EqualTo("10:00"),
                    "COUNTEREXAMPLE: Fissi column 9 'Arrivo' should map to target column 9. " +
                    $"Expected '10:00' in target column 9, but found '{arrivoValue}'. " +
                    "This confirms the bug: Arrivo is being placed in the wrong column.");

                // Requirement 2.6, 2.10: Fissi column 4 "Indirizzo" should NOT be copied to any target column
                // Target column 4 should be empty (it's for assistiti lookup, not fissi data)
                var indirizzoTargetValue = targetWorksheet.Cells[1, 4].Value?.ToString();
                Assert.That(string.IsNullOrEmpty(indirizzoTargetValue), Is.True,
                    "COUNTEREXAMPLE: Fissi column 4 'Indirizzo' should be SKIPPED and not copied to any target column. " +
                    $"Target column 4 should be empty, but found '{indirizzoTargetValue}'. " +
                    "This confirms the bug: Indirizzo is being incorrectly copied.");

                // Requirement 2.7: Target column 10 "Avv" should remain empty
                var avvValue = targetWorksheet.Cells[1, 10].Value?.ToString();
                Assert.That(string.IsNullOrEmpty(avvValue), Is.True,
                    "COUNTEREXAMPLE: Target column 10 'Avv' should remain empty after fissi import. " +
                    $"Expected empty, but found '{avvValue}'. " +
                    "This confirms the bug: Column 10 is being incorrectly populated with fissi column 4 data.");

                // Requirements 2.11, 2.12, 2.13, 2.14, 2.15: Target columns 11-15 should remain empty
                for (int col = 11; col <= 15; col++)
                {
                    var cellValue = targetWorksheet.Cells[1, col].Value?.ToString();
                    Assert.That(string.IsNullOrEmpty(cellValue), Is.True,
                        $"COUNTEREXAMPLE: Target column {col} should remain empty after fissi import. " +
                        $"Expected empty, but found '{cellValue}'. " +
                        "This confirms the bug: Reserved columns are being incorrectly populated.");
                }

                // Verify columns 1-3 are correctly mapped (these should work even on unfixed code)
                Assert.That(targetWorksheet.Cells[1, 1].Value?.ToString(), Is.EqualTo("15/01/2024"),
                    "Fissi column 1 'Data' should map to target column 1");
                Assert.That(targetWorksheet.Cells[1, 2].Value?.ToString(), Is.EqualTo("09:00"),
                    "Fissi column 2 'Partenza' should map to target column 2");
                Assert.That(targetWorksheet.Cells[1, 3].Value?.ToString(), Is.EqualTo("Rossi Mario"),
                    "Fissi column 3 'Assistito' should map to target column 3");
            }
        }

        /// <summary>
        /// Additional exploration test with multiple rows to verify the bug affects all rows.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL for the same reasons as above.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3, 2.4, 2.5, 2.6, 2.7, 2.10, 2.11, 2.12, 2.13, 2.14, 2.15**
        /// </summary>
        [Test]
        public void BugExploration_FissiColumnMapping_MultipleRows_ShouldMapColumnsCorrectly()
        {
            // Arrange - Create a fissi sheet with multiple data rows
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";
                fissiWorksheet.Cells[2, 3].Value = "Assistito";
                fissiWorksheet.Cells[2, 4].Value = "Indirizzo";
                fissiWorksheet.Cells[2, 5].Value = "Destinazione";
                fissiWorksheet.Cells[2, 6].Value = "Note";
                fissiWorksheet.Cells[2, 7].Value = "Auto";
                fissiWorksheet.Cells[2, 8].Value = "Volontario";
                fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                // Add first data row
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 2].Value = "09:00";
                fissiWorksheet.Cells[3, 3].Value = "Rossi Mario";
                fissiWorksheet.Cells[3, 4].Value = "Via Roma 10";
                fissiWorksheet.Cells[3, 5].Value = "Ospedale";
                fissiWorksheet.Cells[3, 6].Value = "Urgente";
                fissiWorksheet.Cells[3, 7].Value = "Auto1";
                fissiWorksheet.Cells[3, 8].Value = "Volontario1";
                fissiWorksheet.Cells[3, 9].Value = "10:00";

                // Add second data row
                fissiWorksheet.Cells[4, 1].Value = "16/01/2024";
                fissiWorksheet.Cells[4, 2].Value = "14:00";
                fissiWorksheet.Cells[4, 3].Value = "Bianchi Luigi";
                fissiWorksheet.Cells[4, 4].Value = "Via Milano 5";
                fissiWorksheet.Cells[4, 5].Value = "Clinica";
                fissiWorksheet.Cells[4, 6].Value = "Normale";
                fissiWorksheet.Cells[4, 7].Value = "Auto2";
                fissiWorksheet.Cells[4, 8].Value = "Volontario2";
                fissiWorksheet.Cells[4, 9].Value = "15:00";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1);

                // Assert - Check both rows
                for (int row = 1; row <= 2; row++)
                {
                    string expectedDestinazione = row == 1 ? "Ospedale" : "Clinica";
                    string expectedNote = row == 1 ? "Urgente" : "Normale";
                    string expectedAuto = row == 1 ? "Auto1" : "Auto2";
                    string expectedVolontario = row == 1 ? "Volontario1" : "Volontario2";
                    string expectedArrivo = row == 1 ? "10:00" : "15:00";

                    // Verify correct mapping for row
                    Assert.That(targetWorksheet.Cells[row, 5].Value?.ToString(), Is.EqualTo(expectedDestinazione),
                        $"Row {row}: Destinazione should be in column 5");
                    Assert.That(targetWorksheet.Cells[row, 6].Value?.ToString(), Is.EqualTo(expectedNote),
                        $"Row {row}: Note should be in column 6");
                    Assert.That(targetWorksheet.Cells[row, 7].Value?.ToString(), Is.EqualTo(expectedAuto),
                        $"Row {row}: Auto should be in column 7");
                    Assert.That(targetWorksheet.Cells[row, 8].Value?.ToString(), Is.EqualTo(expectedVolontario),
                        $"Row {row}: Volontario should be in column 8");
                    Assert.That(targetWorksheet.Cells[row, 9].Value?.ToString(), Is.EqualTo(expectedArrivo),
                        $"Row {row}: Arrivo should be in column 9");

                    // Verify columns 10-15 are empty
                    for (int col = 10; col <= 15; col++)
                    {
                        Assert.That(string.IsNullOrEmpty(targetWorksheet.Cells[row, col].Value?.ToString()), Is.True,
                            $"Row {row}, Column {col} should be empty");
                    }
                }
            }
        }
    }
}
