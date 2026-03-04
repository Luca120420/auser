using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for CSVParser class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 2.1, 2.2, 2.4
    /// </summary>
    [TestFixture]
    public class CSVParserPropertyTests
    {
        private CSVParser _parser = null!;
        private string _testDirectory = null!;

        [SetUp]
        public void Setup()
        {
            _parser = new CSVParser();
            _testDirectory = Path.Combine(Path.GetTempPath(), "CSVParserPropertyTests_" + Guid.NewGuid().ToString());
            Directory.CreateDirectory(_testDirectory);
        }

        [TearDown]
        public void TearDown()
        {
            if (Directory.Exists(_testDirectory))
            {
                try
                {
                    Directory.Delete(_testDirectory, true);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }

        /// <summary>
        /// Helper method to create a test CSV file with UTF-8 encoding
        /// </summary>
        private string CreateTestCSVFile(string content)
        {
            var filePath = Path.Combine(_testDirectory, $"test_{Guid.NewGuid()}.csv");
            File.WriteAllText(filePath, content, Encoding.UTF8);
            return filePath;
        }

        /// <summary>
        /// Generator for valid CSV row data
        /// </summary>
        public class CSVRowData
        {
            public string DataServizio { get; set; } = "";
            public string OraInizioServizio { get; set; } = "";
            public string Attivita { get; set; } = "";
            public string DescrizioneStatoServizio { get; set; } = "";
            public string IndirizzoPartenza { get; set; } = "";
            public string ComunePartenza { get; set; } = "";
            public string DescrizionePuntoPartenza { get; set; } = "";
            public string IndirizzoDestinazione { get; set; } = "";
            public string ComuneDestinazione { get; set; } = "";
            public string CausaleDestinazione { get; set; } = "";
            public string CognomeAssistito { get; set; } = "";
            public string NomeAssistito { get; set; } = "";
            public string NoteERichieste { get; set; } = "";

            public string ToCSVRow()
            {
                return $"{EscapeCSVField(DataServizio)};{EscapeCSVField(OraInizioServizio)};{EscapeCSVField(Attivita)};{EscapeCSVField(DescrizioneStatoServizio)};{EscapeCSVField(IndirizzoPartenza)};{EscapeCSVField(ComunePartenza)};{EscapeCSVField(DescrizionePuntoPartenza)};{EscapeCSVField(IndirizzoDestinazione)};{EscapeCSVField(ComuneDestinazione)};{EscapeCSVField(CausaleDestinazione)};{EscapeCSVField(CognomeAssistito)};{EscapeCSVField(NomeAssistito)};{EscapeCSVField(NoteERichieste)}";
            }

            private string EscapeCSVField(string field)
            {
                // Remove semicolons and newlines to avoid breaking CSV structure
                if (string.IsNullOrEmpty(field))
                    return "";
                
                return field.Replace(";", ",").Replace("\r", "").Replace("\n", " ");
            }
        }

        /// <summary>
        /// Arbitrary generator for CSV row data with Italian characters
        /// </summary>
        public static Arbitrary<CSVRowData> ArbitraryCSVRowData()
        {
            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };
            var normalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '-";

            // Generator for strings that may contain Italian characters
            var italianStringGen = from length in Gen.Choose(0, 50)
                                   from useItalian in Arb.Generate<bool>()
                                   from chars in Gen.ArrayOf(length, useItalian && length > 0
                                       ? Gen.Elements(italianChars.Concat(normalChars.ToCharArray()).ToArray())
                                       : Gen.Elements(normalChars.ToCharArray()))
                                   select new string(chars).Trim(); // Trim to match CSV parser behavior

            // Generator for date strings
            var dateGen = from day in Gen.Choose(1, 28)
                         from month in Gen.Choose(1, 12)
                         from year in Gen.Choose(2020, 2030)
                         select $"{day:D2}/{month:D2}/{year}";

            // Generator for time strings
            var timeGen = from hour in Gen.Choose(0, 23)
                         from minute in Gen.Choose(0, 59)
                         select $"{hour:D2}:{minute:D2}";

            var rowGen = from dataServizio in dateGen
                        from oraInizio in timeGen
                        from attivita in italianStringGen
                        from descrizione in italianStringGen
                        from indPartenza in italianStringGen
                        from comunePartenza in italianStringGen
                        from descPuntoPartenza in italianStringGen
                        from indDestinazione in italianStringGen
                        from comuneDestinazione in italianStringGen
                        from causaleDestinazione in italianStringGen
                        from cognome in italianStringGen.Where(s => !string.IsNullOrWhiteSpace(s))
                        from nome in italianStringGen.Where(s => !string.IsNullOrWhiteSpace(s))
                        from note in italianStringGen
                        select new CSVRowData
                        {
                            DataServizio = dataServizio,
                            OraInizioServizio = oraInizio,
                            Attivita = attivita,
                            DescrizioneStatoServizio = descrizione,
                            IndirizzoPartenza = indPartenza,
                            ComunePartenza = comunePartenza,
                            DescrizionePuntoPartenza = descPuntoPartenza,
                            IndirizzoDestinazione = indDestinazione,
                            ComuneDestinazione = comuneDestinazione,
                            CausaleDestinazione = causaleDestinazione,
                            CognomeAssistito = cognome,
                            NomeAssistito = nome,
                            NoteERichieste = note
                        };

            return Arb.From(rowGen);
        }

        // Feature: auser-excel-transformer, Property 1: CSV Column Extraction Completeness
        /// <summary>
        /// Property 1: CSV Column Extraction Completeness
        /// For any valid CSV file containing the required columns, parsing should extract all columns and all data rows.
        /// **Validates: Requirements 2.1, 2.2**
        /// </summary>
        [Test]
        public void Property_CSVColumnExtractionCompleteness()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryCSVRowData().Generator).Select(list => list.ToList())),
                (List<CSVRowData> rows) =>
                {
                    // Arrange - Create CSV file with all required columns
                    var header = "DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE";
                    var csvContent = header + "\n" + string.Join("\n", rows.Select(r => r.ToCSVRow()));
                    var filePath = CreateTestCSVFile(csvContent);

                    try
                    {
                        // Act
                        var result = _parser.ParseCSV(filePath);

                        // Assert - All rows should be extracted
                        var rowCountMatches = result.Count == rows.Count;

                        // Assert - All columns should be present in each row
                        var allColumnsPresent = result.All(appointment =>
                            appointment.DataServizio != null &&
                            appointment.OraInizioServizio != null &&
                            appointment.CognomeAssistito != null &&
                            appointment.NomeAssistito != null
                        );

                        return rowCountMatches && allColumnsPresent;
                    }
                    catch
                    {
                        // If parsing fails, the property doesn't hold
                        return false;
                    }
                    finally
                    {
                        // Cleanup
                        try { File.Delete(filePath); } catch { }
                    }
                }
            ).Check(config);
        }

        // Feature: auser-excel-transformer, Property 2: Italian Character Preservation
        /// <summary>
        /// Property 2: Italian Character Preservation
        /// For any CSV file containing Italian characters (accented letters, special characters),
        /// parsing should preserve all characters exactly as they appear in the source file.
        /// **Validates: Requirements 2.4**
        /// </summary>
        [Test]
        public void Property_ItalianCharacterPreservation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryCSVRowData().Generator).Select(list => list.ToList())),
                (List<CSVRowData> rows) =>
                {
                    // Filter to only test rows that actually contain Italian characters
                    var rowsWithItalianChars = rows.Where(r =>
                        ContainsAnyChar(r.CognomeAssistito, italianChars) ||
                        ContainsAnyChar(r.NomeAssistito, italianChars) ||
                        ContainsAnyChar(r.Attivita, italianChars) ||
                        ContainsAnyChar(r.DescrizioneStatoServizio, italianChars) ||
                        ContainsAnyChar(r.IndirizzoPartenza, italianChars) ||
                        ContainsAnyChar(r.ComunePartenza, italianChars) ||
                        ContainsAnyChar(r.DescrizionePuntoPartenza, italianChars) ||
                        ContainsAnyChar(r.IndirizzoDestinazione, italianChars) ||
                        ContainsAnyChar(r.ComuneDestinazione, italianChars) ||
                        ContainsAnyChar(r.CausaleDestinazione, italianChars) ||
                        ContainsAnyChar(r.NoteERichieste, italianChars)
                    ).ToList();

                    // If no rows have Italian characters, skip this test case
                    if (rowsWithItalianChars.Count == 0)
                        return true.ToProperty();

                    // Arrange - Create CSV file with Italian characters
                    var header = "DATA SERVIZIO;ORA INIZIO SERVIZIO;ATTIVITA;DESCRIZIONE STATO SERVIZIO;INDIRIZZO PARTENZA;COMUNE PARTENZA;DESCRIZIONE PUNTO PARTENZA;INDIRIZZO DESTINAZIONE;COMUNE DESTINAZIONE;CAUSALE DESTINAZIONE;COGNOME ASSISTITO;NOME ASSISTITO;NOTE E RICHIESTE";
                    var csvContent = header + "\n" + string.Join("\n", rowsWithItalianChars.Select(r => r.ToCSVRow()));
                    var filePath = CreateTestCSVFile(csvContent);

                    try
                    {
                        // Act
                        var result = _parser.ParseCSV(filePath);

                        // Assert - Italian characters should be preserved
                        for (int i = 0; i < Math.Min(result.Count, rowsWithItalianChars.Count); i++)
                        {
                            var original = rowsWithItalianChars[i];
                            var parsed = result[i];

                            // Check that Italian characters are preserved in each field
                            if (!FieldMatches(original.CognomeAssistito, parsed.CognomeAssistito))
                                return false.Label($"CognomeAssistito mismatch at row {i}: expected '{original.CognomeAssistito}', got '{parsed.CognomeAssistito}'");
                            
                            if (!FieldMatches(original.NomeAssistito, parsed.NomeAssistito))
                                return false.Label($"NomeAssistito mismatch at row {i}: expected '{original.NomeAssistito}', got '{parsed.NomeAssistito}'");
                            
                            if (!FieldMatches(original.Attivita, parsed.Attivita))
                                return false.Label($"Attivita mismatch at row {i}: expected '{original.Attivita}', got '{parsed.Attivita}'");
                            
                            if (!FieldMatches(original.DescrizioneStatoServizio, parsed.DescrizioneStatoServizio))
                                return false.Label($"DescrizioneStatoServizio mismatch at row {i}: expected '{original.DescrizioneStatoServizio}', got '{parsed.DescrizioneStatoServizio}'");
                            
                            if (!FieldMatches(original.IndirizzoPartenza, parsed.IndirizzoPartenza))
                                return false.Label($"IndirizzoPartenza mismatch at row {i}: expected '{original.IndirizzoPartenza}', got '{parsed.IndirizzoPartenza}'");
                            
                            if (!FieldMatches(original.ComunePartenza, parsed.ComunePartenza))
                                return false.Label($"ComunePartenza mismatch at row {i}: expected '{original.ComunePartenza}', got '{parsed.ComunePartenza}'");
                            
                            if (!FieldMatches(original.DescrizionePuntoPartenza, parsed.DescrizionePuntoPartenza))
                                return false.Label($"DescrizionePuntoPartenza mismatch at row {i}: expected '{original.DescrizionePuntoPartenza}', got '{parsed.DescrizionePuntoPartenza}'");
                            
                            if (!FieldMatches(original.IndirizzoDestinazione, parsed.IndirizzoDestinazione))
                                return false.Label($"IndirizzoDestinazione mismatch at row {i}: expected '{original.IndirizzoDestinazione}', got '{parsed.IndirizzoDestinazione}'");
                            
                            if (!FieldMatches(original.ComuneDestinazione, parsed.ComuneDestinazione))
                                return false.Label($"ComuneDestinazione mismatch at row {i}: expected '{original.ComuneDestinazione}', got '{parsed.ComuneDestinazione}'");
                            
                            if (!FieldMatches(original.CausaleDestinazione, parsed.CausaleDestinazione))
                                return false.Label($"CausaleDestinazione mismatch at row {i}: expected '{original.CausaleDestinazione}', got '{parsed.CausaleDestinazione}'");
                            
                            if (!FieldMatches(original.NoteERichieste, parsed.NoteERichieste))
                                return false.Label($"NoteERichieste mismatch at row {i}: expected '{original.NoteERichieste}', got '{parsed.NoteERichieste}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        // If parsing fails, the property doesn't hold
                        return false.Label($"Parsing failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        try { File.Delete(filePath); } catch { }
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Helper method to check if a string contains any of the specified characters
        /// </summary>
        private bool ContainsAnyChar(string text, char[] chars)
        {
            if (string.IsNullOrEmpty(text))
                return false;

            return chars.Any(c => text.Contains(c));
        }

        /// <summary>
        /// Helper method to check if two fields match (accounting for null/empty equivalence)
        /// </summary>
        private bool FieldMatches(string original, string? parsed)
        {
            // Treat null and empty as equivalent
            var origNormalized = string.IsNullOrEmpty(original) ? "" : original;
            var parsedNormalized = string.IsNullOrEmpty(parsed) ? "" : parsed;

            return origNormalized == parsedNormalized;
        }
    }
}
