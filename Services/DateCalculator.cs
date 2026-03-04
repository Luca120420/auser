using System;
using System.Collections.Generic;
using System.Globalization;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Provides date calculation and formatting operations with Italian localization.
    /// Implements Requirements 5.3, 5.4, 5.5.
    /// </summary>
    public class DateCalculator : IDateCalculator
    {
        // Italian month abbreviations mapping (1-based index)
        private static readonly Dictionary<int, string> ItalianMonths = new Dictionary<int, string>
        {
            { 1, "gen" },   // gennaio
            { 2, "feb" },   // febbraio
            { 3, "mar" },   // marzo
            { 4, "apr" },   // aprile
            { 5, "mag" },   // maggio
            { 6, "giu" },   // giugno
            { 7, "lug" },   // luglio
            { 8, "ago" },   // agosto
            { 9, "set" },   // settembre
            { 10, "ott" },  // ottobre
            { 11, "nov" },  // novembre
            { 12, "dic" }   // dicembre
        };

        // Reverse mapping for parsing
        private static readonly Dictionary<string, int> MonthNameToNumber = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            { "gen", 1 },
            { "feb", 2 },
            { "mar", 3 },
            { "apr", 4 },
            { "mag", 5 },
            { "giu", 6 },
            { "lug", 7 },
            { "ago", 8 },
            { "set", 9 },
            { "ott", 10 },
            { "nov", 11 },
            { "dic", 12 }
        };

        /// <summary>
        /// Adds a specified number of days to a date.
        /// Handles month and year boundaries correctly.
        /// Implements Requirement 5.3.
        /// </summary>
        public DateTime AddDays(DateTime date, int days)
        {
            return date.AddDays(days);
        }

        /// <summary>
        /// Gets the Italian month abbreviation for a given month number.
        /// Implements Requirement 5.5.
        /// </summary>
        public string GetItalianMonthAbbreviation(int month)
        {
            if (month < 1 || month > 12)
            {
                throw new ArgumentOutOfRangeException(nameof(month), 
                    $"Il mese deve essere compreso tra 1 e 12. Valore ricevuto: {month}");
            }

            return ItalianMonths[month];
        }

        /// <summary>
        /// Formats a date in Italian format with day number and month abbreviation.
        /// Implements Requirement 5.5.
        /// </summary>
        public string FormatItalianDate(DateTime date)
        {
            string monthAbbr = GetItalianMonthAbbreviation(date.Month);
            return $"{date.Day:D2} {monthAbbr}";
        }

        /// <summary>
        /// Parses an Italian date string into a DateTime object.
        /// Expected format: "DD mmm" (e.g., "26 gen", "01 feb")
        /// Implements Requirement 5.1.
        /// </summary>
        public DateTime ParseItalianDate(string dateText, int year)
        {
            if (string.IsNullOrWhiteSpace(dateText))
            {
                throw new FormatException("Il testo della data non può essere vuoto.");
            }

            // Split by space to get day and month parts
            string[] parts = dateText.Trim().Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
            
            if (parts.Length != 2)
            {
                throw new FormatException(
                    $"Formato data non valido: '{dateText}'. Formato atteso: 'DD mmm' (es. '26 gen')");
            }

            // Parse day
            if (!int.TryParse(parts[0], out int day) || day < 1 || day > 31)
            {
                throw new FormatException(
                    $"Giorno non valido: '{parts[0]}'. Deve essere un numero tra 1 e 31.");
            }

            // Parse month abbreviation
            string monthAbbr = parts[1].ToLower();
            if (!MonthNameToNumber.TryGetValue(monthAbbr, out int month))
            {
                throw new FormatException(
                    $"Abbreviazione mese non valida: '{parts[1]}'. " +
                    "Abbreviazioni valide: gen, feb, mar, apr, mag, giu, lug, ago, set, ott, nov, dic");
            }

            try
            {
                return new DateTime(year, month, day);
            }
            catch (ArgumentOutOfRangeException ex)
            {
                throw new FormatException(
                    $"Data non valida: giorno {day}, mese {month}, anno {year}. {ex.Message}", ex);
            }
        }
    }
}
