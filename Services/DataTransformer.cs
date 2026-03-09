using System;
using System.Collections.Generic;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Implements the data transformer service that converts CSV data into Excel format.
    /// Uses the TransformationRulesEngine to apply all transformation rules.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    public class DataTransformer : IDataTransformer
    {
        private readonly ITransformationRulesEngine _rulesEngine;

        /// <summary>
        /// Initializes a new instance of the DataTransformer class.
        /// </summary>
        /// <param name="rulesEngine">The transformation rules engine to use for data transformation</param>
        public DataTransformer(ITransformationRulesEngine rulesEngine)
        {
            _rulesEngine = rulesEngine ?? throw new ArgumentNullException(nameof(rulesEngine));
        }

        /// <summary>
        /// Transforms a list of service appointments into the Excel output format.
        /// Applies all transformation rules and returns the result with highlight information.
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <returns>A TransformationResult containing transformed rows and yellow highlight information</returns>
        /// <exception cref="ArgumentNullException">Thrown when appointments is null</exception>
        public TransformationResult Transform(List<ServiceAppointment> appointments)
        {
            if (appointments == null)
            {
                throw new ArgumentNullException(nameof(appointments));
            }

            // Delegate to the transformation rules engine
            return _rulesEngine.Transform(appointments);
        }

        /// <summary>
        /// Transforms a list of service appointments into the enhanced Excel output format.
        /// Performs lookups against reference sheets and populates all 15 columns.
        /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 7.2, 8.2, 8.3
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <param name="lookupService">The lookup service for reference sheet operations</param>
        /// <returns>An EnhancedTransformationResult containing enhanced rows with lookup data</returns>
        /// <exception cref="ArgumentNullException">Thrown when appointments or lookupService is null</exception>
        public EnhancedTransformationResult TransformEnhanced(List<ServiceAppointment> appointments, ILookupService lookupService)
        {
            if (appointments == null)
            {
                throw new ArgumentNullException(nameof(appointments));
            }

            if (lookupService == null)
            {
                throw new ArgumentNullException(nameof(lookupService));
            }

            var result = new EnhancedTransformationResult();
            int rowIndex = 3; // Row numbering starts at 3 (row 1 = header formulas, row 2 = column headers)

            foreach (var appointment in appointments)
            {
                // Skip cancelled appointments (Rule 3 from original transformation)
                if (IsAnnullato(appointment))
                {
                    continue;
                }

                // Skip rows where all fields are null or empty
                if (IsEmptyRow(appointment))
                {
                    continue;
                }

                // Check if row should be highlighted yellow (Rule 1 from original transformation)
                bool shouldHighlight = ShouldHighlightYellow(appointment);

                // Transform the appointment to an enhanced row with lookups
                var enhancedRow = TransformAppointmentEnhanced(appointment, lookupService);

                result.Rows.Add(enhancedRow);

                // Track rows that need yellow highlighting
                if (shouldHighlight)
                {
                    result.YellowHighlightRows.Add(rowIndex);
                }

                rowIndex++;
            }

            return result;
        }

        /// <summary>
        /// Checks if a row should be highlighted in yellow.
        /// Returns true if ATTIVITÀ contains "Accompag. con macchina attrezzata".
        /// </summary>
        private bool ShouldHighlightYellow(ServiceAppointment appointment)
        {
            return !string.IsNullOrEmpty(appointment.Attivita) &&
                   appointment.Attivita.Contains("Accompag. con macchina attrezzata");
        }

        /// <summary>
        /// Checks if a row should be filtered out.
        /// Returns true if DESCRIZIONE_STATO_SERVIZIO equals "ANNULLATO".
        /// </summary>
        private bool IsAnnullato(ServiceAppointment appointment)
        {
            return !string.IsNullOrEmpty(appointment.DescrizioneStatoServizio) &&
                   appointment.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Checks if all fields in the appointment are null or empty.
        /// Returns true if the row should be excluded from the output.
        /// </summary>
        private bool IsEmptyRow(ServiceAppointment appointment)
        {
            return string.IsNullOrWhiteSpace(appointment.DataServizio) &&
                   string.IsNullOrWhiteSpace(appointment.OraInizioServizio) &&
                   string.IsNullOrWhiteSpace(appointment.Attivita) &&
                   string.IsNullOrWhiteSpace(appointment.DescrizioneStatoServizio) &&
                   string.IsNullOrWhiteSpace(appointment.IndirizzoPartenza) &&
                   string.IsNullOrWhiteSpace(appointment.ComunePartenza) &&
                   string.IsNullOrWhiteSpace(appointment.DescrizionePuntoPartenza) &&
                   string.IsNullOrWhiteSpace(appointment.IndirizzoDestinazione) &&
                   string.IsNullOrWhiteSpace(appointment.ComuneDestinazione) &&
                   string.IsNullOrWhiteSpace(appointment.CausaleDestinazione) &&
                   string.IsNullOrWhiteSpace(appointment.CognomeAssistito) &&
                   string.IsNullOrWhiteSpace(appointment.NomeAssistito) &&
                   string.IsNullOrWhiteSpace(appointment.NoteERichieste);
        }


        /// <summary>
        /// Transforms a single service appointment into an enhanced transformed row.
        /// Performs lookups for Indirizzo, Note (from assistiti), and Avv (from fissi).
        /// Maps CSV Note field to NoteGasnet property.
        /// </summary>
        private EnhancedTransformedRow TransformAppointmentEnhanced(ServiceAppointment appointment, ILookupService lookupService)
        {
            // Create assistito full name for lookups
            string assistitoName = $"{appointment.CognomeAssistito} {appointment.NomeAssistito}";

            // Perform lookups
            string indirizzo = lookupService.LookupInAssistiti(assistitoName, "Indirizzo");
            string note = lookupService.LookupInAssistiti(assistitoName, "Note");
            string avv = lookupService.LookupInFissi(assistitoName, "Avv");

            // Create Destinazione (Rules 11-12 from original transformation)
            string destinazione = CreateDestinazioneColumn(appointment);

            // Create Indirizzo Gasnet (Rule 17 from original transformation)
            string indirizzoGasnet = CreateIndirizzoGasnetColumn(appointment);

            // Create Note Gasnet by concatenating NOTE E RICHIESTE with DESCRIZIONE PUNTO PARTENZA
            string noteGasnet = CreateNoteGasnetColumn(appointment);

            var row = new EnhancedTransformedRow
            {
                // Column 1: Data
                Data = appointment.DataServizio,

                // Column 2: Partenza (renamed from Ora Inizio Servizio)
                Partenza = appointment.OraInizioServizio ?? string.Empty,

                // Column 3: Assistito
                Assistito = assistitoName,

                // Column 4: Indirizzo (from assistiti lookup)
                Indirizzo = indirizzo,

                // Column 5: Destinazione
                Destinazione = destinazione,

                // Column 6: Note (from assistiti lookup)
                Note = note,

                // Column 7: Auto - Empty for CSV
                Auto = string.Empty,

                // Column 8: Volontario - Empty for CSV
                Volontario = string.Empty,

                // Column 9: Arrivo - Empty for CSV
                Arrivo = string.Empty,

                // Column 10: Avv (from fissi lookup)
                Avv = avv,

                // Column 11: Indirizzo Gasnet
                IndirizzoGasnet = indirizzoGasnet,

                // Column 12: Note Gasnet (from CSV Note field + DESCRIZIONE PUNTO PARTENZA)
                NoteGasnet = noteGasnet
            };

            return row;
        }

        /// <summary>
        /// Concatenates DESCRIZIONE PUNTO PARTENZA to INDIRIZZO PARTENZA if present.
        /// </summary>
        private string CreateIndirizzoPartenzaWithDescription(ServiceAppointment appointment)
        {
            var parts = new List<string>();
            
            if (!string.IsNullOrWhiteSpace(appointment.IndirizzoPartenza))
            {
                parts.Add(appointment.IndirizzoPartenza.Trim());
            }
            
            if (!string.IsNullOrWhiteSpace(appointment.DescrizionePuntoPartenza) && 
                !appointment.DescrizionePuntoPartenza.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(appointment.DescrizionePuntoPartenza.Trim());
            }
            
            return string.Join(" ", parts);
        }

        /// <summary>
        /// Creates the DESTINAZIONE column by aggregating 
        /// COMUNE DESTINAZIONE + INDIRIZZO DESTINAZIONE + CAUSALE DESTINAZIONE.
        /// </summary>
        private string CreateDestinazioneColumn(ServiceAppointment appointment)
        {
            var parts = new List<string>();
            
            if (!string.IsNullOrWhiteSpace(appointment.ComuneDestinazione))
            {
                parts.Add(appointment.ComuneDestinazione.Trim());
            }
            
            if (!string.IsNullOrWhiteSpace(appointment.IndirizzoDestinazione))
            {
                parts.Add(appointment.IndirizzoDestinazione.Trim());
            }
            
            if (!string.IsNullOrWhiteSpace(appointment.CausaleDestinazione))
            {
                parts.Add(appointment.CausaleDestinazione.Trim());
            }
            
            return string.Join(" ", parts);
        }

        /// <summary>
        /// Creates the INDIRIZZO GASNET column by aggregating COMUNE PARTENZA + INDIRIZZO PARTENZA.
        /// </summary>
        private string CreateIndirizzoGasnetColumn(ServiceAppointment appointment)
        {
            var parts = new List<string>();
            
            if (!string.IsNullOrWhiteSpace(appointment.ComunePartenza))
            {
                parts.Add(appointment.ComunePartenza.Trim());
            }
            
            if (!string.IsNullOrWhiteSpace(appointment.IndirizzoPartenza))
            {
                parts.Add(appointment.IndirizzoPartenza.Trim());
            }
            
            return string.Join(" ", parts);
        }

        /// <summary>
        /// Creates the NOTE GASNET column by concatenating NOTE E RICHIESTE with DESCRIZIONE PUNTO PARTENZA.
        /// If DESCRIZIONE PUNTO PARTENZA has a value other than null or empty, it is appended to NOTE E RICHIESTE.
        /// </summary>
        private string CreateNoteGasnetColumn(ServiceAppointment appointment)
        {
            var parts = new List<string>();
            
            // Add NOTE E RICHIESTE if present
            if (!string.IsNullOrWhiteSpace(appointment.NoteERichieste))
            {
                parts.Add(appointment.NoteERichieste.Trim());
            }
            
            // Add DESCRIZIONE PUNTO PARTENZA if present and not "null"
            if (!string.IsNullOrWhiteSpace(appointment.DescrizionePuntoPartenza) && 
                !appointment.DescrizionePuntoPartenza.Equals("null", StringComparison.OrdinalIgnoreCase))
            {
                parts.Add(appointment.DescrizionePuntoPartenza.Trim());
            }
            
            return string.Join(" ", parts);
        }
    }
}
