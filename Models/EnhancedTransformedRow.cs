using System;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents an enhanced transformed row with the new 12-column structure.
    /// Columns: Data, Partenza, Assistito, Indirizzo, Destinazione, Note, Auto, Volontario, 
    /// Arrivo, Avv, Indirizzo Gasnet, Note Gasnet
    /// </summary>
    public class EnhancedTransformedRow
    {
        /// <summary>
        /// DATA - The date of the service appointment
        /// </summary>
        [Required(ErrorMessage = "Data è obbligatorio")]
        public string Data { get; set; } = string.Empty;

        /// <summary>
        /// PARTENZA - Departure time (renamed from "Ora Inizio Servizio")
        /// </summary>
        public string Partenza { get; set; } = string.Empty;

        /// <summary>
        /// ASSISTITO - Full name (COGNOME + NOME)
        /// </summary>
        [Required(ErrorMessage = "Assistito è obbligatorio")]
        public string Assistito { get; set; } = string.Empty;

        /// <summary>
        /// INDIRIZZO - Address from assistiti lookup (positioned after Assistito)
        /// </summary>
        public string Indirizzo { get; set; } = string.Empty;

        /// <summary>
        /// DESTINAZIONE - Aggregated destination info
        /// </summary>
        public string Destinazione { get; set; } = string.Empty;

        /// <summary>
        /// NOTE - Notes from assistiti lookup
        /// </summary>
        public string Note { get; set; } = string.Empty;

        /// <summary>
        /// AUTO - Vehicle information
        /// </summary>
        public string Auto { get; set; } = string.Empty;

        /// <summary>
        /// VOLONTARIO - Volunteer name
        /// </summary>
        public string Volontario { get; set; } = string.Empty;

        /// <summary>
        /// ARRIVO - Arrival time
        /// </summary>
        public string Arrivo { get; set; } = string.Empty;

        /// <summary>
        /// AVV - Value from fissi lookup (positioned after Arrivo)
        /// </summary>
        public string Avv { get; set; } = string.Empty;

        /// <summary>
        /// INDIRIZZO GASNET - Aggregated address
        /// </summary>
        public string IndirizzoGasnet { get; set; } = string.Empty;

        /// <summary>
        /// NOTE GASNET - Notes from CSV (positioned after Indirizzo Gasnet)
        /// </summary>
        public string NoteGasnet { get; set; } = string.Empty;
    }
}
