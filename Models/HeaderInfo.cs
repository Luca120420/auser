using System;
using System.ComponentModel.DataAnnotations;

namespace AuserExcelTransformer.Models
{
    /// <summary>
    /// Represents parsed header information from a weekly sheet.
    /// Contains the date range, week number, and referente information.
    /// Validates: Requirements 5.1, 5.2
    /// </summary>
    public class HeaderInfo
    {
        /// <summary>
        /// Monday Date - The Monday date of the week (start of the week range)
        /// Parsed from the header format "DD mmm DD mmm Settimana N..."
        /// </summary>
        public DateTime MondayDate { get; set; }

        /// <summary>
        /// Sunday Date - The Sunday date of the week (end of the week range)
        /// Parsed from the header format "DD mmm DD mmm Settimana N..."
        /// </summary>
        public DateTime SundayDate { get; set; }

        /// <summary>
        /// Week Number - The sequential week number (e.g., Settimana 5)
        /// Extracted from the header text matching "Settimana N"
        /// </summary>
        [Required(ErrorMessage = "WeekNumber è obbligatorio")]
        [Range(1, 53, ErrorMessage = "WeekNumber deve essere tra 1 e 53")]
        public int WeekNumber { get; set; }

        /// <summary>
        /// Referente - The person responsible for coordinating services during the week
        /// Format: "Inserire nome e numero di telefono del referente" for new headers
        /// </summary>
        [Required(ErrorMessage = "Referente è obbligatorio")]
        public string Referente { get; set; } = string.Empty;
    }
}
