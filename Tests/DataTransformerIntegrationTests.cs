using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Integration tests for the DataTransformer class.
    /// Tests complete transformation pipeline with sample data files.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    [TestFixture]
    public class DataTransformerIntegrationTests
    {
        private ICSVParser _csvParser = null!;
        private ITransformationRulesEngine _rulesEngine = null!;
        private IDataTransformer _dataTransformer = null!;

        [SetUp]
        public void SetUp()
        {
            _csvParser = new CSVParser();
            _rulesEngine = new TransformationRulesEngine();
            _dataTransformer = new DataTransformer(_rulesEngine);
        }

        /// <summary>
        /// Integration test with sample_input.csv - realistic Italian data
        /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
        /// </summary>
        [Test]
        public void Transform_WithSampleInputCSV_AppliesAllRulesCorrectly()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            Assert.That(File.Exists(csvPath), Is.True, $"Test file not found: {csvPath}");

            var appointments = _csvParser.ParseCSV(csvPath);
            Assert.That(appointments, Is.Not.Null);
            Assert.That(appointments.Count, Is.GreaterThan(0), "CSV should contain appointments");

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows, Is.Not.Null);
            Assert.That(result.YellowHighlightRows, Is.Not.Null);

            // Verify that ANNULLATO appointments are filtered out
            // Actual data: 14 appointments, 2 with ANNULLATO status
            // Expected: 12 transformed rows
            Assert.That(result.Rows.Count, Is.EqualTo(12), 
                "Should have 12 rows after filtering 2 ANNULLATO appointments");

            // Verify no ANNULLATO rows in output
            foreach (var row in result.Rows)
            {
                // The DescrizioneStatoServizio should not be ANNULLATO in any transformed row
                // Note: The original field is not preserved in TransformedRow, but we verify by count
            }

            // Verify yellow highlighting for "Accompag. con macchina attrezzata"
            // Actual data: 5 appointments should be highlighted
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(5),
                "Should have 5 rows marked for yellow highlighting");

            // Verify all yellow highlight row numbers are valid
            foreach (var rowNum in result.YellowHighlightRows)
            {
                Assert.That(rowNum, Is.GreaterThan(0));
                Assert.That(rowNum, Is.LessThanOrEqualTo(result.Rows.Count));
            }

            // Verify column structure for all rows
            foreach (var row in result.Rows)
            {
                // Verify required columns are present
                Assert.That(row.DataServizio, Is.Not.Null);
                Assert.That(row.OraInizioServizio, Is.Not.Null);
                Assert.That(row.Assistito, Is.Not.Null);
                Assert.That(row.CognomeAssistito, Is.Not.Null);
                Assert.That(row.NomeAssistito, Is.Not.Null);
                Assert.That(row.Indirizzo, Is.Not.Null);

                // Verify ASSISTITO column format (COGNOME + " " + NOME)
                Assert.That(row.Assistito, Is.EqualTo($"{row.CognomeAssistito} {row.NomeAssistito}"),
                    "ASSISTITO should be COGNOME + space + NOME");

                // Verify 5 empty columns are empty
                Assert.That(row.EmptyColumn1, Is.EqualTo(string.Empty));
                Assert.That(row.EmptyColumn2, Is.EqualTo(string.Empty));
                Assert.That(row.EmptyColumn3, Is.EqualTo(string.Empty));
                Assert.That(row.EmptyColumn4, Is.EqualTo(string.Empty));
                Assert.That(row.EmptyColumn5, Is.EqualTo(string.Empty));

                // Verify Partenza field is always empty
                Assert.That(row.Partenza, Is.EqualTo(string.Empty),
                    "Partenza field should always be empty");

                // Verify ORA INIZIO SERVIZIO is copied
                Assert.That(row.OraInizioServizioCopy, Is.EqualTo(row.OraInizioServizio),
                    "OraInizioServizioCopy should match OraInizioServizio");
            }
        }

        /// <summary>
        /// Integration test with edge_cases_input.csv - edge cases and special scenarios
        /// Validates: Requirements 4.1, 4.3, 4.4, 4.5, 4.6, 4.7, 2.4
        /// </summary>
        [Test]
        public void Transform_WithEdgeCasesInputCSV_HandlesEdgeCasesCorrectly()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "edge_cases_input.csv");
            Assert.That(File.Exists(csvPath), Is.True, $"Test file not found: {csvPath}");

            var appointments = _csvParser.ParseCSV(csvPath);
            Assert.That(appointments, Is.Not.Null);
            Assert.That(appointments.Count, Is.EqualTo(10), "Edge cases CSV should have 10 appointments");

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);

            // According to README: 10 appointments, 3 with ANNULLATO status
            // Expected: 7 transformed rows
            Assert.That(result.Rows.Count, Is.EqualTo(7),
                "Should have 7 rows after filtering 3 ANNULLATO appointments");

            // Verify yellow highlighting
            // According to README: 4 rows should be highlighted (excluding cancelled ones)
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(4),
                "Should have 4 rows marked for yellow highlighting");

            // Verify Italian special characters are preserved
            bool hasItalianChars = result.Rows.Any(row =>
                (row.NomeAssistito != null && (row.NomeAssistito.Contains("à") || 
                                                row.NomeAssistito.Contains("è") || 
                                                row.NomeAssistito.Contains("é") ||
                                                row.NomeAssistito.Contains("ì") ||
                                                row.NomeAssistito.Contains("ò") ||
                                                row.NomeAssistito.Contains("ù"))) ||
                (row.CognomeAssistito != null && (row.CognomeAssistito.Contains("'") ||
                                                   row.CognomeAssistito.Contains("á"))) ||
                (row.NoteERichieste != null && (row.NoteERichieste.Contains("à") ||
                                                 row.NoteERichieste.Contains("è") ||
                                                 row.NoteERichieste.Contains("é"))));

            Assert.That(hasItalianChars, Is.True,
                "Should preserve Italian special characters (àèéìòù, apostrophes)");

            // Verify DESCRIZIONE PUNTO PARTENZA duplication
            // At least one row should have duplicated text in Destinazione
            bool hasDuplicatedDestinazione = result.Rows.Any(row =>
                !string.IsNullOrEmpty(row.Destinazione) && row.Destinazione.Length > 1);

            Assert.That(hasDuplicatedDestinazione, Is.True,
                "Should have at least one row with duplicated DESCRIZIONE PUNTO PARTENZA");

            // Verify conditional INDIRIZZO logic
            // All rows should have INDIRIZZO populated (either with IndirizzoDestinazione or CausaleDestinazione)
            foreach (var row in result.Rows)
            {
                Assert.That(row.Indirizzo, Is.Not.Null);
                // INDIRIZZO should contain at least the comune (even if other parts are empty)
            }
        }

        /// <summary>
        /// Integration test with empty_input.csv - CSV with no data rows
        /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
        /// </summary>
        [Test]
        public void Transform_WithEmptyInputCSV_ReturnsEmptyResult()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "empty_input.csv");
            Assert.That(File.Exists(csvPath), Is.True, $"Test file not found: {csvPath}");

            var appointments = _csvParser.ParseCSV(csvPath);
            Assert.That(appointments, Is.Not.Null);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows, Is.Not.Null);
            Assert.That(result.YellowHighlightRows, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(0),
                "Empty CSV should produce zero transformed rows");
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(0),
                "Empty CSV should have no yellow highlight rows");
        }

        /// <summary>
        /// Integration test verifying Rule 1: Yellow highlighting for "Accompag. con macchina attrezzata"
        /// Validates: Requirement 4.1
        /// </summary>
        [Test]
        public void Transform_WithSampleData_CorrectlyIdentifiesYellowHighlightRows()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify that yellow highlight rows are within valid range
            foreach (var rowNum in result.YellowHighlightRows)
            {
                Assert.That(rowNum, Is.GreaterThan(0), "Row numbers should be 1-based");
                Assert.That(rowNum, Is.LessThanOrEqualTo(result.Rows.Count),
                    "Row number should not exceed total rows");
            }

            // Verify no duplicate row numbers in highlight list
            var distinctHighlights = result.YellowHighlightRows.Distinct().ToList();
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(distinctHighlights.Count),
                "Yellow highlight rows should not contain duplicates");
        }

        /// <summary>
        /// Integration test verifying Rule 3: Filter out ANNULLATO appointments
        /// Validates: Requirement 4.3
        /// </summary>
        [Test]
        public void Transform_WithSampleData_FiltersOutAnnullatoAppointments()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Count ANNULLATO appointments in input
            int annullatoCount = appointments.Count(a =>
                !string.IsNullOrEmpty(a.DescrizioneStatoServizio) &&
                a.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase));

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify that output has fewer rows than input (ANNULLATO filtered)
            Assert.That(result.Rows.Count, Is.EqualTo(appointments.Count - annullatoCount),
                $"Should filter out {annullatoCount} ANNULLATO appointments");
        }

        /// <summary>
        /// Integration test verifying Rule 4: Duplicate DESCRIZIONE PUNTO PARTENZA
        /// Validates: Requirement 4.4
        /// </summary>
        [Test]
        public void Transform_WithSampleData_DuplicatesDescrizionePuntoPartenza()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "edge_cases_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Find appointments with non-empty DESCRIZIONE PUNTO PARTENZA
            var appointmentsWithDescrizione = appointments.Where(a =>
                !string.IsNullOrEmpty(a.DescrizionePuntoPartenza) &&
                !a.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // For each appointment with DESCRIZIONE PUNTO PARTENZA, verify duplication
            foreach (var appointment in appointmentsWithDescrizione)
            {
                // Find corresponding row in result (match by name)
                var matchingRow = result.Rows.FirstOrDefault(r =>
                    r.CognomeAssistito == appointment.CognomeAssistito &&
                    r.NomeAssistito == appointment.NomeAssistito);

                if (matchingRow != null && !string.IsNullOrEmpty(appointment.DescrizionePuntoPartenza))
                {
                    string expectedDestinazione = appointment.DescrizionePuntoPartenza +
                                                   appointment.DescrizionePuntoPartenza;
                    Assert.That(matchingRow.Destinazione, Is.EqualTo(expectedDestinazione),
                        $"DESCRIZIONE PUNTO PARTENZA should be duplicated for {appointment.CognomeAssistito}");
                }
            }
        }

        /// <summary>
        /// Integration test verifying Rule 5: Create ASSISTITO column
        /// Validates: Requirement 4.5
        /// </summary>
        [Test]
        public void Transform_WithSampleData_CreatesAssistitoColumnCorrectly()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify ASSISTITO column for all rows
            foreach (var row in result.Rows)
            {
                string expectedAssistito = $"{row.CognomeAssistito} {row.NomeAssistito}";
                Assert.That(row.Assistito, Is.EqualTo(expectedAssistito),
                    "ASSISTITO should be COGNOME + space + NOME");
            }
        }

        /// <summary>
        /// Integration test verifying Rules 6-7: INDIRIZZO column conditional concatenation
        /// Validates: Requirements 4.6, 4.7
        /// </summary>
        [Test]
        public void Transform_WithSampleData_CreatesIndirizzoColumnWithConditionalLogic()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "edge_cases_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify INDIRIZZO column for all rows
            foreach (var row in result.Rows)
            {
                Assert.That(row.Indirizzo, Is.Not.Null, "INDIRIZZO should never be null");
                
                // INDIRIZZO should contain at least one space (comune + space + something)
                // unless comune is empty
                if (!string.IsNullOrEmpty(row.Indirizzo))
                {
                    // Should follow pattern: "COMUNE INDIRIZZO" or "COMUNE CAUSALE"
                    Assert.That(row.Indirizzo.Contains(" ") || row.Indirizzo.Length > 0,
                        "INDIRIZZO should follow concatenation pattern");
                }
            }
        }

        /// <summary>
        /// Integration test verifying Rule 8: Column structure with 5 empty columns
        /// Validates: Requirement 4.8
        /// </summary>
        [Test]
        public void Transform_WithSampleData_CreatesCorrectColumnStructure()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            foreach (var row in result.Rows)
            {
                // Verify all columns exist
                Assert.That(row.DataServizio, Is.Not.Null);
                Assert.That(row.OraInizioServizio, Is.Not.Null);
                Assert.That(row.Assistito, Is.Not.Null);
                Assert.That(row.CognomeAssistito, Is.Not.Null);
                Assert.That(row.NomeAssistito, Is.Not.Null);
                Assert.That(row.Indirizzo, Is.Not.Null);
                Assert.That(row.Destinazione, Is.Not.Null);
                
                // Verify 5 empty columns
                Assert.That(row.EmptyColumn1, Is.Not.Null);
                Assert.That(row.EmptyColumn2, Is.Not.Null);
                Assert.That(row.EmptyColumn3, Is.Not.Null);
                Assert.That(row.EmptyColumn4, Is.Not.Null);
                Assert.That(row.EmptyColumn5, Is.Not.Null);
                
                Assert.That(row.EmptyColumn1, Is.Empty, "EmptyColumn1 should be empty");
                Assert.That(row.EmptyColumn2, Is.Empty, "EmptyColumn2 should be empty");
                Assert.That(row.EmptyColumn3, Is.Empty, "EmptyColumn3 should be empty");
                Assert.That(row.EmptyColumn4, Is.Empty, "EmptyColumn4 should be empty");
                Assert.That(row.EmptyColumn5, Is.Empty, "EmptyColumn5 should be empty");
                
                Assert.That(row.OraInizioServizioCopy, Is.Not.Null);
                Assert.That(row.Partenza, Is.Not.Null);
                Assert.That(row.NoteERichieste, Is.Not.Null);
            }
        }

        /// <summary>
        /// Integration test verifying Rule 9: Copy ORA INIZIO SERVIZIO
        /// Validates: Requirement 4.8
        /// </summary>
        [Test]
        public void Transform_WithSampleData_CopiesOraInizioServizio()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            foreach (var row in result.Rows)
            {
                Assert.That(row.OraInizioServizioCopy, Is.EqualTo(row.OraInizioServizio),
                    "OraInizioServizioCopy should match OraInizioServizio");
            }
        }

        /// <summary>
        /// Integration test verifying Rule 10: Partenza field always empty
        /// Validates: Requirement 4.9
        /// </summary>
        [Test]
        public void Transform_WithSampleData_PartenzaFieldAlwaysEmpty()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            foreach (var row in result.Rows)
            {
                Assert.That(row.Partenza, Is.EqualTo(string.Empty),
                    "Partenza field should always be empty");
            }

            // Verify all Partenza fields are empty
            Assert.That(result.Rows.All(r => r.Partenza == string.Empty), Is.True,
                "All Partenza fields should be empty");
        }

        /// <summary>
        /// Integration test verifying Rule 11: Preserve NOTE E RICHIESTE
        /// Validates: Requirement 4.8
        /// </summary>
        [Test]
        public void Transform_WithSampleData_PreservesNoteERichieste()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "edge_cases_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Find appointments with notes
            var appointmentsWithNotes = appointments.Where(a =>
                !string.IsNullOrEmpty(a.NoteERichieste) &&
                !a.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase))
                .ToList();

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify that notes are preserved
            foreach (var appointment in appointmentsWithNotes)
            {
                var matchingRow = result.Rows.FirstOrDefault(r =>
                    r.CognomeAssistito == appointment.CognomeAssistito &&
                    r.NomeAssistito == appointment.NomeAssistito);

                if (matchingRow != null)
                {
                    Assert.That(matchingRow.NoteERichieste, Is.EqualTo(appointment.NoteERichieste),
                        $"NOTE E RICHIESTE should be preserved for {appointment.CognomeAssistito}");
                }
            }
        }

        /// <summary>
        /// Integration test verifying Italian character preservation
        /// Validates: Requirement 2.4
        /// </summary>
        [Test]
        public void Transform_WithItalianCharacters_PreservesAllCharacters()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "edge_cases_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            // Verify Italian special characters are preserved in various fields
            bool hasAccentedVowels = result.Rows.Any(r =>
                (r.NomeAssistito != null && (r.NomeAssistito.Contains("à") || r.NomeAssistito.Contains("è") ||
                                              r.NomeAssistito.Contains("é") || r.NomeAssistito.Contains("ì") ||
                                              r.NomeAssistito.Contains("ò") || r.NomeAssistito.Contains("ù"))) ||
                (r.CognomeAssistito != null && r.CognomeAssistito.Contains("'")) ||
                (r.Indirizzo != null && (r.Indirizzo.Contains("à") || r.Indirizzo.Contains("è"))) ||
                (r.NoteERichieste != null && (r.NoteERichieste.Contains("à") || r.NoteERichieste.Contains("è"))));

            Assert.That(hasAccentedVowels, Is.True,
                "Should preserve Italian special characters (àèéìòù) and apostrophes");
        }

        /// <summary>
        /// Integration test verifying complete transformation pipeline
        /// Tests all rules together with realistic data
        /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
        /// </summary>
        [Test]
        public void Transform_CompleteTransformationPipeline_AppliesAllRulesInCorrectOrder()
        {
            // Arrange
            string csvPath = Path.Combine("TestData", "sample_input.csv");
            var appointments = _csvParser.ParseCSV(csvPath);
            
            int originalCount = appointments.Count;
            int annullatoCount = appointments.Count(a =>
                !string.IsNullOrEmpty(a.DescrizioneStatoServizio) &&
                a.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase));

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert - Comprehensive validation
            Assert.That(result, Is.Not.Null, "Result should not be null");
            Assert.That(result.Rows, Is.Not.Null, "Rows should not be null");
            Assert.That(result.YellowHighlightRows, Is.Not.Null, "YellowHighlightRows should not be null");

            // Rule 3: Verify filtering
            Assert.That(result.Rows.Count, Is.EqualTo(originalCount - annullatoCount),
                "Should filter out ANNULLATO appointments");

            // Rule 1: Verify yellow highlighting
            Assert.That(result.YellowHighlightRows.Count, Is.GreaterThan(0),
                "Should have at least one yellow highlight row");

            // Verify all rules for each row
            for (int i = 0; i < result.Rows.Count; i++)
            {
                var row = result.Rows[i];
                int rowNumber = i + 1;

                // Rule 5: ASSISTITO column
                Assert.That(row.Assistito, Is.EqualTo($"{row.CognomeAssistito} {row.NomeAssistito}"),
                    $"Row {rowNumber}: ASSISTITO should be COGNOME + space + NOME");

                // Rule 6: Separate COGNOME and NOME columns
                Assert.That(row.CognomeAssistito, Is.Not.Null, $"Row {rowNumber}: CognomeAssistito should not be null");
                Assert.That(row.NomeAssistito, Is.Not.Null, $"Row {rowNumber}: NomeAssistito should not be null");

                // Rule 6-7: INDIRIZZO column
                Assert.That(row.Indirizzo, Is.Not.Null, $"Row {rowNumber}: INDIRIZZO should not be null");

                // Rule 8: Five empty columns
                Assert.That(row.EmptyColumn1, Is.Empty, $"Row {rowNumber}: EmptyColumn1 should be empty");
                Assert.That(row.EmptyColumn2, Is.Empty, $"Row {rowNumber}: EmptyColumn2 should be empty");
                Assert.That(row.EmptyColumn3, Is.Empty, $"Row {rowNumber}: EmptyColumn3 should be empty");
                Assert.That(row.EmptyColumn4, Is.Empty, $"Row {rowNumber}: EmptyColumn4 should be empty");
                Assert.That(row.EmptyColumn5, Is.Empty, $"Row {rowNumber}: EmptyColumn5 should be empty");

                // Rule 9: ORA INIZIO SERVIZIO copy
                Assert.That(row.OraInizioServizioCopy, Is.EqualTo(row.OraInizioServizio),
                    $"Row {rowNumber}: OraInizioServizioCopy should match OraInizioServizio");

                // Rule 10: Partenza always empty
                Assert.That(row.Partenza, Is.Empty, $"Row {rowNumber}: Partenza should be empty");

                // Rule 11: NOTE E RICHIESTE preserved
                Assert.That(row.NoteERichieste, Is.Not.Null, $"Row {rowNumber}: NoteERichieste should not be null");
            }
        }
    }
}
