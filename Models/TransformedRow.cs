using System;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents a transformed row in the output Excel sheet with the new column structure.
    /// Columns: DATA SERVIZIO, ORA INIZIO SERVIZIO (empty for CSV), ASSISTITO, DESTINAZIONE, NOTE, AUTO, VOLONTARIO, 
    /// ARRIVO, [empty], INDIRIZZO PARTENZA, COMUNE PARTENZA, [empty], [empty], INDIRIZZO GASNET
    /// </summary>
    public class TransformedRow
    {
        /// <summary>
        /// DATA SERVIZIO - The date of the service appointment
        /// </summary>
        [Required(ErrorMessage = "Data è obbligatorio")]
        public string DataServizio { get; set; } = string.Empty;

        /// <summary>
        /// ORA INIZIO SERVIZIO - Empty for CSV, has value for fissi
        /// </summary>
        public string OraInizioServizio { get; set; } = string.Empty;

        /// <summary>
        /// ASSISTITO - Full name (COGNOME + NOME)
        /// </summary>
        [Required(ErrorMessage = "Assistito è obbligatorio")]
        public string Assistito { get; set; } = string.Empty;

        /// <summary>
        /// DESTINAZIONE - Aggregated destination info
        /// </summary>
        public string Destinazione { get; set; } = string.Empty;

        /// <summary>
        /// NOTE - Notes and requests
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
        /// ARRIVO - Arrival time (from ORA INIZIO SERVIZIO)
        /// </summary>
        public string Arrivo { get; set; } = string.Empty;

        /// <summary>
        /// Empty column 1
        /// </summary>
        public string Empty1 { get; set; } = string.Empty;

        /// <summary>
        /// INDIRIZZO PARTENZA - Departure address (with DESCRIZIONE PUNTO PARTENZA if present)
        /// </summary>
        public string IndirizzoPartenza { get; set; } = string.Empty;

        /// <summary>
        /// COMUNE PARTENZA - Departure municipality
        /// </summary>
        public string ComunePartenza { get; set; } = string.Empty;

        /// <summary>
        /// Empty column 2
        /// </summary>
        public string Empty2 { get; set; } = string.Empty;

        /// <summary>
        /// Empty column 3
        /// </summary>
        public string Empty3 { get; set; } = string.Empty;

        /// <summary>
        /// INDIRIZZO GASNET - Aggregated address (COMUNE PARTENZA + INDIRIZZO PARTENZA)
        /// </summary>
        public string IndirizzoGasnet { get; set; } = string.Empty;
    }
}
