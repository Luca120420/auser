using System;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents a service appointment record from the CSV file.
    /// Contains all 13 columns from the CSV export.
    /// Validates: Requirements 2.2
    /// </summary>
    public class ServiceAppointment
    {
        /// <summary>
        /// DATA SERVIZIO - The date of the service appointment
        /// </summary>
        [Required(ErrorMessage = "DATA SERVIZIO è obbligatorio")]
        public string DataServizio { get; set; } = string.Empty;

        /// <summary>
        /// ORA INIZIO SERVIZIO - The start time of the service
        /// </summary>
        [Required(ErrorMessage = "ORA INIZIO SERVIZIO è obbligatorio")]
        public string OraInizioServizio { get; set; } = string.Empty;

        /// <summary>
        /// ATTIVITÀ - The activity type (e.g., "Accompag. con macchina attrezzata")
        /// </summary>
        public string? Attivita { get; set; }

        /// <summary>
        /// DESCRIZIONE STATO SERVIZIO - The service status description (e.g., "ANNULLATO")
        /// </summary>
        public string? DescrizioneStatoServizio { get; set; }

        /// <summary>
        /// INDIRIZZO PARTENZA - The departure address
        /// </summary>
        public string? IndirizzoPartenza { get; set; }

        /// <summary>
        /// COMUNE PARTENZA - The departure municipality
        /// </summary>
        public string? ComunePartenza { get; set; }

        /// <summary>
        /// DESCRIZIONE PUNTO PARTENZA - The departure point description
        /// </summary>
        public string? DescrizionePuntoPartenza { get; set; }

        /// <summary>
        /// INDIRIZZO DESTINAZIONE - The destination address
        /// </summary>
        public string? IndirizzoDestinazione { get; set; }

        /// <summary>
        /// COMUNE DESTINAZIONE - The destination municipality
        /// </summary>
        public string? ComuneDestinazione { get; set; }

        /// <summary>
        /// CAUSALE DESTINAZIONE - The destination reason/purpose
        /// </summary>
        public string? CausaleDestinazione { get; set; }

        /// <summary>
        /// COGNOME ASSISTITO - The last name of the elderly person receiving care
        /// </summary>
        [Required(ErrorMessage = "COGNOME ASSISTITO è obbligatorio")]
        public string CognomeAssistito { get; set; } = string.Empty;

        /// <summary>
        /// NOME ASSISTITO - The first name of the elderly person receiving care
        /// </summary>
        [Required(ErrorMessage = "NOME ASSISTITO è obbligatorio")]
        public string NomeAssistito { get; set; } = string.Empty;

        /// <summary>
        /// NOTE E RICHIESTE - Notes and requests for the service
        /// </summary>
        public string? NoteERichieste { get; set; }
    }
}
