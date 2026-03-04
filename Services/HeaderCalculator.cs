using System;
using System.Text.RegularExpressions;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Provides header calculation operations for weekly sheets.
    /// Parses existing headers and generates new week headers with proper date arithmetic.
    /// Implements Requirements 5.1, 5.2, 5.6, 5.7, 5.8.
    /// </summary>
    public class HeaderCalculator : IHeaderCalculator
    {
        private readonly IDateCalculator _dateCalculator;
        
        // Default referente text for new headers (Requirement 5.7)
        private const string DefaultReferenteText = "Inserire nome e numero di telefono del referente";

        // Regex pattern to parse header format: "DD mmm DD mmm Settimana N..."
        // Example: "26 gen 01 feb Settimana 5referente settimana = Inserire nome e numero di telefono del referente"
        private static readonly Regex HeaderPattern = new Regex(
            @"^(\d{1,2})\s+(\w+)\s+(\d{1,2})\s+(\w+)\s+Settimana\s+(\d+)",
            RegexOptions.IgnoreCase | RegexOptions.Compiled
        );

        /// <summary>
        /// Initializes a new instance of the HeaderCalculator class.
        /// </summary>
        /// <param name="dateCalculator">The date calculator service for date operations.</param>
        public HeaderCalculator(IDateCalculator dateCalculator)
        {
            _dateCalculator = dateCalculator ?? throw new ArgumentNullException(nameof(dateCalculator));
        }

        /// <summary>
        /// Parses a header string to extract date range, week number, and referente information.
        /// Implements Requirements 5.1, 5.2.
        /// </summary>
        /// <param name="headerText">The header text to parse.</param>
        /// <returns>A HeaderInfo object containing the parsed information.</returns>
        /// <exception cref="FormatException">Thrown when the header text cannot be parsed.</exception>
        public HeaderInfo ParseHeader(string headerText)
        {
            if (string.IsNullOrWhiteSpace(headerText))
            {
                throw new FormatException("Il testo dell'intestazione non può essere vuoto.");
            }

            // Match the header pattern
            Match match = HeaderPattern.Match(headerText.Trim());
            
            if (!match.Success)
            {
                throw new FormatException(
                    $"Formato intestazione non valido: '{headerText}'. " +
                    "Formato atteso: 'DD mmm DD mmm Settimana N...' (es. '26 gen 01 feb Settimana 5...')");
            }

            // Extract matched groups
            string mondayDay = match.Groups[1].Value;
            string mondayMonth = match.Groups[2].Value;
            string sundayDay = match.Groups[3].Value;
            string sundayMonth = match.Groups[4].Value;
            string weekNumberStr = match.Groups[5].Value;

            // Parse week number
            if (!int.TryParse(weekNumberStr, out int weekNumber) || weekNumber < 1 || weekNumber > 53)
            {
                throw new FormatException(
                    $"Numero settimana non valido: '{weekNumberStr}'. Deve essere tra 1 e 53.");
            }

            // Determine the year for date parsing
            // We need to handle year boundaries (e.g., week spanning Dec-Jan)
            int currentYear = DateTime.Now.Year;
            
            // Parse Monday date
            DateTime mondayDate;
            try
            {
                mondayDate = _dateCalculator.ParseItalianDate($"{mondayDay} {mondayMonth}", currentYear);
            }
            catch (FormatException ex)
            {
                throw new FormatException(
                    $"Errore nell'analisi della data di lunedì: '{mondayDay} {mondayMonth}'. {ex.Message}", ex);
            }

            // Parse Sunday date
            DateTime sundayDate;
            try
            {
                // Sunday might be in the next year if we're at year boundary
                int sundayYear = currentYear;
                
                // If Monday is in December and Sunday month is January, Sunday is next year
                if (mondayDate.Month == 12 && mondayMonth.ToLower() == "dic" && 
                    (sundayMonth.ToLower() == "gen" || sundayMonth.ToLower() == "gennaio"))
                {
                    sundayYear = currentYear + 1;
                }
                
                sundayDate = _dateCalculator.ParseItalianDate($"{sundayDay} {sundayMonth}", sundayYear);
                
                // If Sunday is before Monday, it must be in the next year
                if (sundayDate < mondayDate)
                {
                    sundayDate = _dateCalculator.ParseItalianDate($"{sundayDay} {sundayMonth}", currentYear + 1);
                }
            }
            catch (FormatException ex)
            {
                throw new FormatException(
                    $"Errore nell'analisi della data di domenica: '{sundayDay} {sundayMonth}'. {ex.Message}", ex);
            }

            // Extract referente text (everything after "referente settimana = ")
            string referente = DefaultReferenteText;
            int referenteIndex = headerText.IndexOf("referente settimana = ", StringComparison.OrdinalIgnoreCase);
            if (referenteIndex >= 0)
            {
                referente = headerText.Substring(referenteIndex + "referente settimana = ".Length).Trim();
            }

            return new HeaderInfo
            {
                MondayDate = mondayDate,
                SundayDate = sundayDate,
                WeekNumber = weekNumber,
                Referente = referente
            };
        }

        /// <summary>
        /// Generates the next week's header based on the previous week's header.
        /// Implements Requirements 5.3, 5.4, 5.5, 5.6, 5.7, 5.8.
        /// </summary>
        /// <param name="previousHeader">The previous week's header text.</param>
        /// <returns>The formatted header text for the next week.</returns>
        /// <exception cref="FormatException">Thrown when the previous header cannot be parsed.</exception>
        public string GenerateNextWeekHeader(string previousHeader)
        {
            // Parse the previous header
            HeaderInfo previousHeaderInfo = ParseHeader(previousHeader);

            // Calculate next week's dates (add 7 days) - Requirement 5.3, 5.4
            DateTime nextMonday = _dateCalculator.AddDays(previousHeaderInfo.MondayDate, 7);
            DateTime nextSunday = _dateCalculator.AddDays(previousHeaderInfo.SundayDate, 7);

            // Increment week number - Requirement 5.6
            int nextWeekNumber = previousHeaderInfo.WeekNumber + 1;

            // Format dates using Italian month abbreviations - Requirement 5.5
            string mondayFormatted = _dateCalculator.FormatItalianDate(nextMonday);
            string sundayFormatted = _dateCalculator.FormatItalianDate(nextSunday);

            // Generate header in the required format - Requirement 5.8
            // Format: "DD mmm DD mmm Settimana Nreferente settimana = Inserire nome e numero di telefono del referente"
            string newHeader = $"{mondayFormatted} {sundayFormatted} Settimana {nextWeekNumber}referente settimana = {DefaultReferenteText}";

            return newHeader;
        }
    }
}
