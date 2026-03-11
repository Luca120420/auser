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
        /// Static constructor to register encoding provider for code page encodings
        /// </summary>
        static CSVParser()
        {
            // Register the code page encoding provider to support encodings like Windows-1252
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }

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

            // Try multiple encodings to handle different CSV sources
            var encodingsToTry = new[]
            {
                new UTF8Encoding(true), // UTF-8 with BOM (most likely for Italian CSV files)
                Encoding.UTF8, // UTF-8 without BOM
                Encoding.GetEncoding("ISO-8859-1"), // Latin-1, common for Italian files
                Encoding.GetEncoding(1252), // Windows-1252, Western European
                Encoding.Default // System default encoding
            };

            Exception lastException = null;

            foreach (var encoding in encodingsToTry)
            {
                try
                {
                    // Configure CsvHelper with the current encoding
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        Encoding = encoding,
                        Delimiter = ";", // Italian CSV files use semicolon delimiter
                        HasHeaderRecord = true,
                        TrimOptions = TrimOptions.Trim,
                        MissingFieldFound = null, // Don't throw on missing fields
                        BadDataFound = null // Handle bad data gracefully
                    };

                    using (var reader = new StreamReader(filePath, encoding, true)) // detectEncodingFromByteOrderMarks = true
                    using (var csv = new CsvReader(reader, config))
                    {
                        // Register the class map for ServiceAppointment
                        csv.Context.RegisterClassMap<ServiceAppointmentMap>();

                        // Read all records
                        appointments = csv.GetRecords<ServiceAppointment>().ToList();
                        
                        // If we successfully read records, return them
                        if (appointments.Count > 0)
                        {
                            return appointments;
                        }
                    }
                }
                catch (Exception ex)
                {
                    lastException = ex;
                    // Try next encoding
                    continue;
                }
            }

            // If all encodings failed, throw the last exception
            if (lastException != null)
            {
                if (lastException is CsvHelper.CsvHelperException)
                {
                    throw new IOException($"Errore durante la lettura del file CSV: {lastException.Message}", lastException);
                }
                throw new IOException($"Impossibile leggere il file CSV: {lastException.Message}", lastException);
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

            // Try multiple encodings to handle different CSV sources
            var encodingsToTry = new[]
            {
                new UTF8Encoding(true), // UTF-8 with BOM (most likely for Italian CSV files)
                Encoding.UTF8, // UTF-8 without BOM
                Encoding.GetEncoding("ISO-8859-1"), // Latin-1, common for Italian files
                Encoding.GetEncoding(1252), // Windows-1252, Western European
                Encoding.Default // System default encoding
            };

            foreach (var encoding in encodingsToTry)
            {
                try
                {
                    var config = new CsvConfiguration(CultureInfo.InvariantCulture)
                    {
                        Encoding = encoding,
                        Delimiter = ";", // Italian CSV files use semicolon delimiter
                        HasHeaderRecord = true
                    };

                    using (var reader = new StreamReader(filePath, encoding, true)) // detectEncodingFromByteOrderMarks = true
                    using (var csv = new CsvReader(reader, config))
                    {
                        // Read the header
                        csv.Read();
                        csv.ReadHeader();

                        // Get all header names
                        var headers = csv.HeaderRecord;

                        if (headers == null)
                        {
                            // Try next encoding
                            continue;
                        }

                        // Check each required column and track which ones are missing
                        var tempMissingColumns = new List<string>();
                        foreach (var requiredColumn in RequiredColumns)
                        {
                            if (!headers.Contains(requiredColumn, StringComparer.OrdinalIgnoreCase))
                            {
                                tempMissingColumns.Add(requiredColumn);
                            }
                        }

                        // If we found all columns with this encoding, return success
                        if (tempMissingColumns.Count == 0)
                        {
                            missingColumns = tempMissingColumns;
                            return true;
                        }

                        // Keep track of the best result (fewest missing columns)
                        if (missingColumns.Count == 0 || tempMissingColumns.Count < missingColumns.Count)
                        {
                            missingColumns = tempMissingColumns;
                        }
                    }
                }
                catch
                {
                    // Try next encoding
                    continue;
                }
            }

            // If we get here, we didn't find all columns with any encoding
            // Return false with the best result we found
            if (missingColumns.Count == 0)
            {
                // No headers found with any encoding - add all columns as missing
                missingColumns.AddRange(RequiredColumns);
            }
            
            return false;
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
