using System;
using System.Collections.Generic;
using System.Linq;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Preprocesses CSV data according to transformation rules before Excel export.
    /// Restructures columns and applies business logic transformations.
    /// </summary>
    public class CSVPreprocessor
    {
        /// <summary>
        /// Preprocesses a list of service appointments according to transformation rules.
        /// Returns a list of preprocessed rows ready for Excel export.
        /// </summary>
        public List<PreprocessedRow> Preprocess(List<ServiceAppointment> appointments)
        {
            if (appointments == null)
            {
                throw new ArgumentNullException(nameof(appointments));
            }

            var result = new List<PreprocessedRow>();

            foreach (var appointment in appointments)
            {
                // Rule 3: Skip rows with "ANNULLATO"
                if (!string.IsNullOrEmpty(appointment.DescrizioneStatoServizio) &&
                    appointment.DescrizioneStatoServizio.Equals("ANNULLATO", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                // Rule 1: Check if should highlight (Accompag. con macchina attrezzata)
                bool shouldHighlight = !string.IsNullOrEmpty(appointment.Attivita) &&
                                      appointment.Attivita.Contains("Accompag. con macchina attrezzata");

                // Create preprocessed row
                var row = new PreprocessedRow
                {
                    DataServizio = appointment.DataServizio,
                    OraInizioServizio = string.Empty, // Rule 16: Clear content
                    Assistito = CreateAssistito(appointment), // Rule 8
                    Indirizzo = string.Empty, // Rule 10: Empty column
                    Destinazione = CreateDestinazione(appointment), // Rule 12
                    Empty1 = string.Empty, // Rule 14: 5 empty columns
                    Empty2 = string.Empty,
                    Empty3 = string.Empty,
                    Arrivo = appointment.OraInizioServizio, // Rule 15: 4th column after DESTINAZIONE
                    Empty5 = string.Empty,
                    IndirizzoPartenza = CreateIndirizzoPartenza(appointment), // Rule 5
                    ComunePartenza = appointment.ComunePartenza ?? string.Empty,
                    NoteERichieste = appointment.NoteERichieste ?? string.Empty,
                    IndirizzoGasnet = CreateIndirizzoGasnet(appointment), // Rule 17
                    ShouldHighlight = shouldHighlight
                };

                result.Add(row);
            }

            return result;
        }

        /// <summary>
        /// Rule 8: Concatenate COGNOME ASSISTITO and NOME ASSISTITO with space
        /// </summary>
        private string CreateAssistito(ServiceAppointment appointment)
        {
            return $"{appointment.CognomeAssistito} {appointment.NomeAssistito}".Trim();
        }

        /// <summary>
        /// Rule 12: Concatenate COMUNE DESTINAZIONE + INDIRIZZO DESTINAZIONE + CAUSALE DESTINAZIONE
        /// </summary>
        private string CreateDestinazione(ServiceAppointment appointment)
        {
            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(appointment.ComuneDestinazione))
                parts.Add(appointment.ComuneDestinazione.Trim());

            if (!string.IsNullOrWhiteSpace(appointment.IndirizzoDestinazione))
                parts.Add(appointment.IndirizzoDestinazione.Trim());

            if (!string.IsNullOrWhiteSpace(appointment.CausaleDestinazione))
                parts.Add(appointment.CausaleDestinazione.Trim());

            return string.Join(" ", parts);
        }

        /// <summary>
        /// Rule 5: If DESCRIZIONE PUNTO PARTENZA has content, concatenate to INDIRIZZO PARTENZA
        /// </summary>
        private string CreateIndirizzoPartenza(ServiceAppointment appointment)
        {
            string indirizzo = appointment.IndirizzoPartenza ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(appointment.DescrizionePuntoPartenza))
            {
                indirizzo = (indirizzo + " " + appointment.DescrizionePuntoPartenza).Trim();
            }

            return indirizzo;
        }

        /// <summary>
        /// Rule 17: Concatenate COMUNE PARTENZA and INDIRIZZO PARTENZA
        /// </summary>
        private string CreateIndirizzoGasnet(ServiceAppointment appointment)
        {
            var parts = new List<string>();

            if (!string.IsNullOrWhiteSpace(appointment.ComunePartenza))
                parts.Add(appointment.ComunePartenza.Trim());

            if (!string.IsNullOrWhiteSpace(appointment.IndirizzoPartenza))
                parts.Add(appointment.IndirizzoPartenza.Trim());

            return string.Join(" ", parts);
        }
    }

    /// <summary>
    /// Represents a preprocessed row with the new column structure
    /// </summary>
    public class PreprocessedRow
    {
        public string DataServizio { get; set; } = string.Empty;
        public string OraInizioServizio { get; set; } = string.Empty; // Empty per rule 16
        public string Assistito { get; set; } = string.Empty;
        public string Indirizzo { get; set; } = string.Empty;
        public string Destinazione { get; set; } = string.Empty;
        public string Empty1 { get; set; } = string.Empty;
        public string Empty2 { get; set; } = string.Empty;
        public string Empty3 { get; set; } = string.Empty;
        public string Arrivo { get; set; } = string.Empty;
        public string Empty5 { get; set; } = string.Empty;
        public string IndirizzoPartenza { get; set; } = string.Empty;
        public string ComunePartenza { get; set; } = string.Empty;
        public string NoteERichieste { get; set; } = string.Empty;
        public string IndirizzoGasnet { get; set; } = string.Empty;
        public bool ShouldHighlight { get; set; }
    }
}
