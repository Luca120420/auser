using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using CsvHelper;
using CsvHelper.Configuration;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// CSV parsing service implementation using CsvHelper library.
    /// Handles UTF-8 encoding for Italian character support.
    /// Validates: Requirements 2.1, 2.2, 2.4
    /// </summary>
    public class CSVParser : ICSVParser
    {
        /// <summary>
        /// Required CSV column names as they appear in the source file
        /// Note: The actual CSV uses "ATTIVITA" without accent
        /// </summary>
        private static readonly string[] RequiredColumns = new[]
        {
            "DATA SERVIZIO",
            "ORA INIZIO SERVIZIO",
            "ATTIVITA", // Note: No accent in actual CSV
            "DESCRIZIONE STATO SERVIZIO",
            "INDIRIZZO PARTENZA",
            "COMUNE PARTENZA",
            "DESCRIZIONE PUNTO PARTENZA",
            "INDIRIZZO DESTINAZIONE",
            "COMUNE DESTINAZIONE",
            "CAUSALE DESTINAZIONE",
            "COGNOME ASSISTITO",
            "NOME ASSISTITO",
            "NOTE E RICHIESTE"
        };

        /// <summary>
        /// Parses a CSV file and returns a list of ServiceAppointment objects.
        /// Uses UTF-8 encoding to preserve Italian characters.
        /// </summary>
        /// <param name="filePath">The path to the CSV file to parse</param>
        /// <returns>A list of ServiceAppointment objects parsed from the CSV file</returns>
        /// <exception cref="FileNotFoundException">Thrown when the CSV file is not found</exception>
        /// <exception cref="IOException">Thrown when the CSV file cannot be read</exception>
        /// <exception cref="CsvHelper.CsvHelperException">Thrown when the CSV file is malformed</exception>
        public List<ServiceAppointment> ParseCSV(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Il file CSV non è stato trovato: {filePath}", filePath);
            }

            var appointments = new List<ServiceAppointment>();

            try
            {
                // Configure CsvHelper to use UTF-8 encoding for Italian character support
                // Italian CSV files typically use semicolon as delimiter
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Encoding = Encoding.UTF8,
                    Delimiter = ";", // Italian CSV files use semicolon delimiter
                    HasHeaderRecord = true,
                    TrimOptions = TrimOptions.Trim,
                    MissingFieldFound = null, // Don't throw on missing fields
                    BadDataFound = null // Handle bad data gracefully
                };

                using (var reader = new StreamReader(filePath, Encoding.UTF8))
                using (var csv = new CsvReader(reader, config))
                {
                    // Register the class map for ServiceAppointment
                    csv.Context.RegisterClassMap<ServiceAppointmentMap>();

                    // Read all records
                    appointments = csv.GetRecords<ServiceAppointment>().ToList();
                }
            }
            catch (CsvHelper.CsvHelperException ex)
            {
                throw new IOException($"Errore durante la lettura del file CSV: {ex.Message}", ex);
            }
            catch (Exception ex)
            {
                throw new IOException($"Impossibile leggere il file CSV: {ex.Message}", ex);
            }

            return appointments;
        }

        /// <summary>
        /// Validates that a CSV file contains all required columns.
        /// </summary>
        /// <param name="filePath">The path to the CSV file to validate</param>
        /// <returns>True if the CSV file has all required columns, false otherwise</returns>
        public bool ValidateCSVStructure(string filePath)
        {
            List<string> missingColumns;
            return ValidateCSVStructure(filePath, out missingColumns);
        }

        /// <summary>
        /// Validates that a CSV file contains all required columns and returns detailed error information.
        /// Validates: Requirements 2.3, 9.3
        /// </summary>
        /// <param name="filePath">The path to the CSV file to validate</param>
        /// <param name="missingColumns">Output parameter containing the list of missing column names</param>
        /// <returns>True if the CSV file has all required columns, false otherwise</returns>
        public bool ValidateCSVStructure(string filePath, out List<string> missingColumns)
        {
            missingColumns = new List<string>();

            if (!File.Exists(filePath))
            {
                // File doesn't exist - add all columns as missing
                missingColumns.AddRange(RequiredColumns);
                return false;
            }

            try
            {
                var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                {
                    Encoding = Encoding.UTF8,
                    Delimiter = ";", // Italian CSV files use semicolon delimiter
                    HasHeaderRecord = true
                };

                using (var reader = new StreamReader(filePath, Encoding.UTF8))
                using (var csv = new CsvReader(reader, config))
                {
                    // Read the header
                    csv.Read();
                    csv.ReadHeader();

                    // Get all header names
                    var headers = csv.HeaderRecord;

                    if (headers == null)
                    {
                        // No headers found - add all columns as missing
                        missingColumns.AddRange(RequiredColumns);
                        return false;
                    }

                    // Check each required column and track which ones are missing
                    foreach (var requiredColumn in RequiredColumns)
                    {
                        if (!headers.Contains(requiredColumn, StringComparer.OrdinalIgnoreCase))
                        {
                            missingColumns.Add(requiredColumn);
                        }
                    }

                    return missingColumns.Count == 0;
                }
            }
            catch
            {
                // On any error, consider all columns as potentially missing
                missingColumns.AddRange(RequiredColumns);
                return false;
            }
        }
    }

    /// <summary>
    /// CsvHelper class map for mapping CSV columns to ServiceAppointment properties.
    /// Handles the exact column names from the CSV file.
    /// </summary>
    public sealed class ServiceAppointmentMap : ClassMap<ServiceAppointment>
    {
        public ServiceAppointmentMap()
        {
            Map(m => m.DataServizio).Name("DATA SERVIZIO");
            Map(m => m.OraInizioServizio).Name("ORA INIZIO SERVIZIO");
            Map(m => m.Attivita).Name("ATTIVITA").Optional(); // Note: No accent in actual CSV
            Map(m => m.DescrizioneStatoServizio).Name("DESCRIZIONE STATO SERVIZIO").Optional();
            Map(m => m.IndirizzoPartenza).Name("INDIRIZZO PARTENZA").Optional();
            Map(m => m.ComunePartenza).Name("COMUNE PARTENZA").Optional();
            Map(m => m.DescrizionePuntoPartenza).Name("DESCRIZIONE PUNTO PARTENZA").Optional();
            Map(m => m.IndirizzoDestinazione).Name("INDIRIZZO DESTINAZIONE").Optional();
            Map(m => m.ComuneDestinazione).Name("COMUNE DESTINAZIONE").Optional();
            Map(m => m.CausaleDestinazione).Name("CAUSALE DESTINAZIONE").Optional();
            Map(m => m.CognomeAssistito).Name("COGNOME ASSISTITO");
            Map(m => m.NomeAssistito).Name("NOME ASSISTITO");
            Map(m => m.NoteERichieste).Name("NOTE E RICHIESTE").Optional();
        }
    }
}
