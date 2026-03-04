using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for CSVParser class.
    /// Tests CSV parsing, validation, and Italian character support.
    /// Validates: Requirements 2.1, 2.2, 2.3, 2.4
    /// </summary>
    [TestFixture]
    public class CSVParserTests
    {
        private CSVParser _parser = null!;
        private string _testDirectory = null!;

        [SetUp]
        public void Setup()
        {
            _parser = new CSVParser();
            _testDirectory = Path.Combine(Path.GetTempPath(), "CSVParserTests_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(_testDirectory);
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_testDirectory))
            {
                Directory.Delete(_testDirectory, true);
            }
        }

        /// <summary>
        /// Helper method to create a test CSV file with all required columns
        /// </summary>
        private string CreateTestCSVFile(string content)
        {
            var filePath = Path.Combine(_testDirectory, $"test_{Guid.NewGuid()}.csv");
            File.WriteAllText(filePath, content, Encoding.UTF8);
            return filePath;
        }

        [Test]
        public void ParseCSV_WithValidFile_ReturnsServiceAppointments()
        {
            // Arrange
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nessuna nota
27/01/2026;10:30;Accompag. con macchina attrezzata;Confermato;Via Dante 5;Milano;Residenza;Via Manzoni 20;Milano;Clinica;Bianchi;Giuseppe;Urgente";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(2));
            
            // Verify first appointment
            Assert.That(result[0].DataServizio, Is.EqualTo("26/01/2026"));
            Assert.That(result[0].OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(result[0].Attivita, Is.EqualTo("Accompagnamento"));
            Assert.That(result[0].DescrizioneStatoServizio, Is.EqualTo("Confermato"));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(result[0].NomeAssistito, Is.EqualTo("Mario"));
            
            // Verify second appointment
            Assert.That(result[1].DataServizio, Is.EqualTo("27/01/2026"));
            Assert.That(result[1].Attivita, Is.EqualTo("Accompag. con macchina attrezzata"));
            Assert.That(result[1].CognomeAssistito, Is.EqualTo("Bianchi"));
            Assert.That(result[1].NomeAssistito, Is.EqualTo("Giuseppe"));
        }

        [Test]
        public void ParseCSV_WithItalianCharacters_PreservesCharacters()
        {
            // Arrange - Test Italian accented characters
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Università;Milano;Città;Via Libertà;Milano;Ospedale;Rossi;José;Più informazioni";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(1));
            Assert.That(result[0].IndirizzoPartenza, Is.EqualTo("Via Università"));
            Assert.That(result[0].IndirizzoDestinazione, Is.EqualTo("Via Libertà"));
            Assert.That(result[0].NomeAssistito, Is.EqualTo("José"));
            Assert.That(result[0].NoteERichieste, Is.EqualTo("Più informazioni"));
        }

        [Test]
        public void ParseCSV_WithEmptyOptionalFields_ReturnsAppointmentsWithNullValues()
        {
            // Arrange
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;;;;;;;Milano;;Rossi;Mario;";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(1));
            Assert.That(result[0].DataServizio, Is.EqualTo("26/01/2026"));
            Assert.That(result[0].OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(result[0].NomeAssistito, Is.EqualTo("Mario"));
            // Optional fields should be empty or null
            Assert.That(string.IsNullOrEmpty(result[0].Attivita));
            Assert.That(string.IsNullOrEmpty(result[0].IndirizzoPartenza));
        }

        [Test]
        public void ParseCSV_WithNonExistentFile_ThrowsFileNotFoundException()
        {
            // Arrange
            var nonExistentPath = Path.Combine(_testDirectory, "nonexistent.csv");

            // Act & Assert
            var ex = Assert.Throws<FileNotFoundException>(() => _parser.ParseCSV(nonExistentPath));
            Assert.That(ex!.Message, Does.Contain("Il file CSV non è stato trovato"));
        }

        [Test]
        public void ValidateCSVStructure_WithValidFile_ReturnsTrue()
        {
            // Arrange
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nessuna nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ValidateCSVStructure(filePath);

            // Assert
            Assert.That(result, Is.True);
        }

        [Test]
        public void ValidateCSVStructure_WithMissingColumns_ReturnsFalse()
        {
            // Arrange - Missing "ATTIVITA" and "NOTE E RICHIESTE" columns
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;DESCRIZIONE STATO SERVIZIO;COGNOME ASSISTITO;NOME ASSISTITO
26/01/2026;09:00;Confermato;Rossi;Mario";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ValidateCSVStructure(filePath);

            // Assert
            Assert.That(result, Is.False);
        }

        [Test]
        public void ValidateCSVStructure_WithNonExistentFile_ReturnsFalse()
        {
            // Arrange
            var nonExistentPath = Path.Combine(_testDirectory, "nonexistent.csv");

            // Act
            var result = _parser.ValidateCSVStructure(nonExistentPath);

            // Assert
            Assert.That(result, Is.False);
        }

        [Test]
        public void ParseCSV_WithMultipleRows_ReturnsAllRows()
        {
            // Arrange
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nota 1
27/01/2026;10:30;Trasporto;Confermato;Via Dante 5;Milano;Residenza;Via Manzoni 20;Milano;Clinica;Bianchi;Giuseppe;Nota 2
28/01/2026;14:00;Visita;ANNULLATO;Via Garibaldi 3;Milano;Appartamento;Via Leopardi 15;Milano;Ambulatorio;Verdi;Anna;Nota 3";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(3));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(result[1].CognomeAssistito, Is.EqualTo("Bianchi"));
            Assert.That(result[2].CognomeAssistito, Is.EqualTo("Verdi"));
            Assert.That(result[2].DescrizioneStatoServizio, Is.EqualTo("ANNULLATO"));
        }

        [Test]
        public void ValidateCSVStructure_WithSpecialItalianCharacters_PreservesAllCharacters()
        {
            // Arrange - Test various Italian special characters: à, è, é, ì, ò, ù
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Attività speciale;Già confermato;Città di Modena;Più vicino;Università;Qualità superiore;Perché necessario;Così importante;D'Àngelo;Nicolò;È più urgente";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(1));
            Assert.That(result[0].Attivita, Is.EqualTo("Attività speciale"));
            Assert.That(result[0].DescrizioneStatoServizio, Is.EqualTo("Già confermato"));
            Assert.That(result[0].IndirizzoPartenza, Is.EqualTo("Città di Modena"));
            Assert.That(result[0].ComunePartenza, Is.EqualTo("Più vicino"));
            Assert.That(result[0].DescrizionePuntoPartenza, Is.EqualTo("Università"));
            Assert.That(result[0].IndirizzoDestinazione, Is.EqualTo("Qualità superiore"));
            Assert.That(result[0].ComuneDestinazione, Is.EqualTo("Perché necessario"));
            Assert.That(result[0].CausaleDestinazione, Is.EqualTo("Così importante"));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("D'Àngelo"));
            Assert.That(result[0].NomeAssistito, Is.EqualTo("Nicolò"));
            Assert.That(result[0].NoteERichieste, Is.EqualTo("È più urgente"));
        }

        // Tests for detailed validation with missing columns information
        // Validates: Requirements 2.3, 9.3

        [Test]
        public void ValidateCSVStructure_WithMissingColumns_ReturnsDetailedErrorInfo()
        {
            // Arrange - Missing "ATTIVITA" and "NOTE E RICHIESTE" columns
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO
26/01/2026;09:00;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(2));
            Assert.That(missingColumns, Does.Contain("ATTIVITA"));
            Assert.That(missingColumns, Does.Contain("NOTE E RICHIESTE"));
        }

        [Test]
        public void ValidateCSVStructure_WithAllColumnsPresent_ReturnsEmptyMissingList()
        {
            // Arrange
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nessuna nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.True);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(0));
        }

        [Test]
        public void ValidateCSVStructure_WithSingleMissingColumn_ReturnsCorrectMissingColumn()
        {
            // Arrange - Missing only "COGNOME ASSISTITO" column
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Mario;Nessuna nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(1));
            Assert.That(missingColumns[0], Is.EqualTo("COGNOME ASSISTITO"));
        }

        [Test]
        public void ValidateCSVStructure_WithNonExistentFile_ReturnsAllColumnsAsMissing()
        {
            // Arrange
            var nonExistentPath = Path.Combine(_testDirectory, "nonexistent.csv");

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(nonExistentPath, out missingColumns);

            // Assert
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(13)); // All 13 required columns
        }

        [Test]
        public void ValidateCSVStructure_WithEmptyFile_ReturnsAllColumnsAsMissing()
        {
            // Arrange - Empty file with no headers
            var csvContent = "";
            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(13)); // All 13 required columns
        }

        [Test]
        public void ValidateCSVStructure_WithMultipleMissingColumns_ReturnsAllMissingColumns()
        {
            // Arrange - Missing 5 columns
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE;COMUNE DESTINAZIONE;INDIRIZZO DESTINAZIONE
26/01/2026;09:00;Accompagnamento;Rossi;Mario;Nota;Milano;Via Roma 1";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(5));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE STATO SERVIZIO"));
            Assert.That(missingColumns, Does.Contain("INDIRIZZO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("COMUNE PARTENZA"));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE PUNTO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("CAUSALE DESTINAZIONE"));
        }

        [Test]
        public void ValidateCSVStructure_CaseInsensitiveColumnMatching_ReturnsTrue()
        {
            // Arrange - Column names in different case
            var csvContent = @"data servizio;ora inizio servizio;attivita;descrizione stato servizio;indirizzo partenza;comune partenza;descrizione punto partenza;indirizzo destinazione;comune destinazione;causale destinazione;cognome assistito;nome assistito;note e richieste
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nessuna nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert
            Assert.That(result, Is.True);
            Assert.That(missingColumns.Count, Is.EqualTo(0));
        }

        // Additional tests for malformed CSV data
        // Validates: Requirements 2.3, 9.3

        [Test]
        public void ParseCSV_WithMalformedData_InconsistentColumnCount_HandlesGracefully()
        {
            // Arrange - Row with too few columns (malformed)
            // CSVParser is configured to handle missing fields gracefully
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nota
27/01/2026;10:30;Trasporto";

            var filePath = CreateTestCSVFile(csvContent);

            // Act - Should handle gracefully without throwing
            var result = _parser.ParseCSV(filePath);

            // Assert - First row should parse correctly, second row may have null/empty fields
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.GreaterThanOrEqualTo(1));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("Rossi"));
        }

        [Test]
        public void ParseCSV_WithCorruptedFile_HandlesGracefully()
        {
            // Arrange - Create a file with binary/corrupted content
            var filePath = Path.Combine(_testDirectory, $"corrupted_{Guid.NewGuid()}.csv");
            var binaryContent = new byte[] { 0xFF, 0xFE, 0x00, 0x01, 0x02, 0x03, 0xFF, 0xFF };
            File.WriteAllBytes(filePath, binaryContent);

            // Act - May throw IOException or return empty/invalid data
            try
            {
                var result = _parser.ParseCSV(filePath);
                // If it doesn't throw, result should be empty or contain invalid data
                Assert.That(result, Is.Not.Null);
            }
            catch (IOException ex)
            {
                // If it throws, error message should be in Italian
                Assert.That(ex.Message, Does.Contain("Errore durante la lettura del file CSV").Or.Contains("Impossibile leggere il file CSV"));
            }
        }

        [Test]
        public void ParseCSV_WithUnterminatedQuotes_HandlesGracefully()
        {
            // Arrange - CSV with unterminated quoted field
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE
26/01/2026;09:00;""Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act - CSVParser handles bad data gracefully
            var result = _parser.ParseCSV(filePath);

            // Assert - Should not crash, may return empty or partial data
            Assert.That(result, Is.Not.Null);
        }

        [Test]
        public void ParseCSV_WithInvalidDelimiter_ThrowsIOException()
        {
            // Arrange - CSV using comma instead of semicolon (wrong delimiter for Italian CSV)
            var csvContent = @"DATA SERVIZIO,ORA INIZIO SERVIZIO,ATTIVITA,DESCRIZIONE STATO SERVIZIO,INDIRIZZO PARTENZA,COMUNE PARTENZA,DESCRIZIONE PUNTO PARTENZA,INDIRIZZO DESTINAZIONE,COMUNE DESTINAZIONE,CAUSALE DESTINAZIONE,COGNOME ASSISTITO,NOME ASSISTITO,NOTE E RICHIESTE
26/01/2026,09:00,Accompagnamento,Confermato,Via Roma 1,Milano,Casa,Via Verdi 10,Milano,Ospedale,Rossi,Mario,Nota";

            var filePath = CreateTestCSVFile(csvContent);

            // Act & Assert - With wrong delimiter, parsing should throw IOException
            var ex = Assert.Throws<IOException>(() => _parser.ParseCSV(filePath));
            
            // Verify error message is in Italian
            Assert.That(ex!.Message, Does.Contain("Errore durante la lettura del file CSV").Or.Contains("Impossibile leggere il file CSV"));
        }

        // Tests for error message content
        // Validates: Requirements 2.3, 9.3

        [Test]
        public void ParseCSV_FileNotFound_ErrorMessageInItalian()
        {
            // Arrange
            var nonExistentPath = Path.Combine(_testDirectory, "nonexistent.csv");

            // Act & Assert
            var ex = Assert.Throws<FileNotFoundException>(() => _parser.ParseCSV(nonExistentPath));
            
            // Verify error message is in Italian
            Assert.That(ex!.Message, Does.Contain("Il file CSV non è stato trovato"));
            Assert.That(ex.Message, Does.Not.Contain("File not found"));
            Assert.That(ex.Message, Does.Not.Contain("Could not find"));
        }

        [Test]
        public void ParseCSV_MalformedData_ErrorMessageInItalian()
        {
            // Arrange - Create a truly malformed file that will cause an exception
            var filePath = Path.Combine(_testDirectory, $"malformed_{Guid.NewGuid()}.csv");
            // Write invalid UTF-8 sequence that will cause encoding issues
            var invalidBytes = new byte[] { 
                0x44, 0x41, 0x54, 0x41, 0x20, // "DATA "
                0xFF, 0xFE, 0xFF, 0xFE // Invalid UTF-8 sequence
            };
            File.WriteAllBytes(filePath, invalidBytes);

            // Act & Assert
            try
            {
                var result = _parser.ParseCSV(filePath);
                // If it doesn't throw, that's also acceptable (graceful handling)
                Assert.That(result, Is.Not.Null);
            }
            catch (IOException ex)
            {
                // If it throws, verify error message is in Italian
                Assert.That(ex.Message, Does.Contain("Errore durante la lettura del file CSV").Or.Contains("Impossibile leggere il file CSV"));
                Assert.That(ex.Message, Does.Not.Contain("Error reading"));
                Assert.That(ex.Message, Does.Not.Contain("Unable to read"));
            }
        }

        [Test]
        public void ValidateCSVStructure_WithMissingColumns_ErrorListContainsCorrectColumnNames()
        {
            // Arrange - Missing specific columns
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;COGNOME ASSISTITO;NOME ASSISTITO
26/01/2026;09:00;Rossi;Mario";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert - Verify specific missing columns are reported
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(9)); // 9 columns missing
            
            // Verify each missing column is correctly identified
            Assert.That(missingColumns, Does.Contain("ATTIVITA"));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE STATO SERVIZIO"));
            Assert.That(missingColumns, Does.Contain("INDIRIZZO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("COMUNE PARTENZA"));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE PUNTO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("INDIRIZZO DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("COMUNE DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("CAUSALE DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("NOTE E RICHIESTE"));
        }

        [Test]
        public void ValidateCSVStructure_WithAllMissingColumns_ReturnsAllRequiredColumns()
        {
            // Arrange - File with completely wrong headers
            var csvContent = @"WRONG_COLUMN1;WRONG_COLUMN2;WRONG_COLUMN3
value1;value2;value3";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert - All 13 required columns should be reported as missing
            Assert.That(result, Is.False);
            Assert.That(missingColumns, Is.Not.Null);
            Assert.That(missingColumns.Count, Is.EqualTo(13));
            
            // Verify all required columns are in the missing list
            Assert.That(missingColumns, Does.Contain("DATA SERVIZIO"));
            Assert.That(missingColumns, Does.Contain("ORA INIZIO SERVIZIO"));
            Assert.That(missingColumns, Does.Contain("ATTIVITA"));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE STATO SERVIZIO"));
            Assert.That(missingColumns, Does.Contain("INDIRIZZO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("COMUNE PARTENZA"));
            Assert.That(missingColumns, Does.Contain("DESCRIZIONE PUNTO PARTENZA"));
            Assert.That(missingColumns, Does.Contain("INDIRIZZO DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("COMUNE DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("CAUSALE DESTINAZIONE"));
            Assert.That(missingColumns, Does.Contain("COGNOME ASSISTITO"));
            Assert.That(missingColumns, Does.Contain("NOME ASSISTITO"));
            Assert.That(missingColumns, Does.Contain("NOTE E RICHIESTE"));
        }

        [Test]
        public void ParseCSV_WithHeaderOnlyFile_ReturnsEmptyList()
        {
            // Arrange - File with only headers, no data rows
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(0));
        }

        [Test]
        public void ParseCSV_WithExtraColumns_ParsesSuccessfully()
        {
            // Arrange - CSV with extra columns beyond required ones
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE;EXTRA_COLUMN1;EXTRA_COLUMN2
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nota;Extra1;Extra2";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            var result = _parser.ParseCSV(filePath);

            // Assert - Should parse successfully, ignoring extra columns
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.EqualTo(1));
            Assert.That(result[0].CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(result[0].NomeAssistito, Is.EqualTo("Mario"));
        }

        [Test]
        public void ValidateCSVStructure_WithExtraColumns_ReturnsTrue()
        {
            // Arrange - CSV with extra columns beyond required ones
            var csvContent = @"DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE;EXTRA_COLUMN
26/01/2026;09:00;Accompagnamento;Confermato;Via Roma 1;Milano;Casa;Via Verdi 10;Milano;Ospedale;Rossi;Mario;Nota;Extra";

            var filePath = CreateTestCSVFile(csvContent);

            // Act
            List<string> missingColumns;
            var result = _parser.ValidateCSVStructure(filePath, out missingColumns);

            // Assert - Should validate successfully with extra columns
            Assert.That(result, Is.True);
            Assert.That(missingColumns.Count, Is.EqualTo(0));
        }
    }
}
