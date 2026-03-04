using System;
using System.Collections.Generic;
using System.Linq;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Implements the transformation rules engine that applies all 17 data transformation rules.
    /// New structure: DATA SERVIZIO, ORA INIZIO SERVIZIO (empty), ASSISTITO, INDIRIZZO (empty), 
    /// DESTINAZIONE, [3 empty], ARRIVO, [1 empty], INDIRIZZO PARTENZA, COMUNE PARTENZA, NOTE, INDIRIZZO GASNET
    /// </summary>
    public class TransformationRulesEngine : ITransformationRulesEngine
    {
        /// <summary>
        /// Transforms a list of service appointments according to all 17 transformation rules.
        /// </summary>
        /// <param name="appointments">The list of service appointments from the CSV file</param>
        /// <returns>A TransformationResult containing transformed rows and yellow highlight information</returns>
        public TransformationResult Transform(List<ServiceAppointment> appointments)
        {
            if (appointments == null)
            {
                throw new ArgumentNullException(nameof(appointments));
            }

            var result = new TransformationResult();
            int rowIndex = 3; // Row numbering starts at 3 (row 1 = header formulas, row 2 = column headers)

            foreach (var appointment in appointments)
            {
                // Rule 3: Delete rows with "ANNULLATO"
                if (IsAnnullato(appointment))
                {
                    continue;
                }

                // Rule 1: Identify rows with "Accompag. con macchina attrezzata" for yellow highlighting
                bool shouldHighlight = ShouldHighlightYellow(appointment);

                // Transform the appointment to a row
                var transformedRow = TransformAppointment(appointment);

                result.Rows.Add(transformedRow);

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
        /// Rule 1: Identifies if a row should be highlighted in yellow.
        /// Checks if ATTIVITÀ contains "Accompag. con macchina attrezzata".
        /// </summary>
        private bool ShouldHighlightYellow(ServiceAppointment appointment)
        {
            return !string.IsNullOrEmpty(appointment.Attivita) &&
                   appointment.Attivita.Contains("Accompag. con macchina attrezzata");
        }

        /// <summary>
        /// Rule 3: Checks if a row should be filtered out.
        /// Returns true if DESCRIZIONE_STATO_SERVIZIO equals "ANNULLATO".
        /// </summary>
        private bool IsAnnullato(ServiceAppointment appointment)
        {
            return !string.IsNullOrEmpty(appointment.DescrizioneStatoServizio) &&
                   appointment.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Transforms a single service appointment into a transformed row.
        /// Implements all 17 transformation rules.
        /// </summary>
        private TransformedRow TransformAppointment(ServiceAppointment appointment)
        {
            // Rule 5: Concatenate DESCRIZIONE PUNTO PARTENZA to INDIRIZZO PARTENZA if present
            string indirizzoPartenza = CreateIndirizzoPartenzaWithDescription(appointment);

            var row = new TransformedRow
            {
                // Column 1: DATA SERVIZIO
                DataServizio = appointment.DataServizio,

                // Column 2: ORA INIZIO SERVIZIO - From CSV ORA INIZIO SERVIZIO
                OraInizioServizio = appointment.OraInizioServizio ?? string.Empty,

                // Column 3: ASSISTITO - Rule 7-8: COGNOME + " " + NOME
                Assistito = CreateAssistitoName(appointment),

                // Column 4: DESTINAZIONE - Rule 11-12: COMUNE DESTINAZIONE + INDIRIZZO DESTINAZIONE + CAUSALE DESTINAZIONE
                Destinazione = CreateDestinazioneColumn(appointment),

                // Column 5: NOTE - From NOTE E RICHIESTE
                Note = appointment.NoteERichieste ?? string.Empty,

                // Column 6: AUTO - Empty for CSV
                Auto = string.Empty,

                // Column 7: VOLONTARIO - Empty for CSV
                Volontario = string.Empty,

                // Column 8: ARRIVO - Empty for CSV
                Arrivo = string.Empty,

                // Column 9: Empty column
                Empty1 = string.Empty,

                // Column 10: INDIRIZZO PARTENZA - Rule 5: With DESCRIZIONE PUNTO PARTENZA concatenated
                IndirizzoPartenza = indirizzoPartenza,

                // Column 11: COMUNE PARTENZA
                ComunePartenza = appointment.ComunePartenza ?? string.Empty,

                // Columns 12-13: Empty columns
                Empty2 = string.Empty,
                Empty3 = string.Empty,

                // Column 14: INDIRIZZO GASNET - Rule 17: COMUNE PARTENZA + INDIRIZZO PARTENZA
                IndirizzoGasnet = CreateIndirizzoGasnetColumn(appointment)
            };

            return row;
        }

        /// <summary>
        /// Rule 5: Concatenates DESCRIZIONE PUNTO PARTENZA to INDIRIZZO PARTENZA if present.
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
        /// Rule 7-8: Creates the ASSISTITO column by concatenating COGNOME + " " + NOME.
        /// </summary>
        private string CreateAssistitoName(ServiceAppointment appointment)
        {
            return $"{appointment.CognomeAssistito} {appointment.NomeAssistito}";
        }

        /// <summary>
        /// Rule 11-12: Creates the DESTINAZIONE column by aggregating 
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
        /// Rule 17: Creates the INDIRIZZO GASNET column by aggregating COMUNE PARTENZA + INDIRIZZO PARTENZA.
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
    }
}
