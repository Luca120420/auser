using System;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for date calculation and formatting operations with Italian localization.
    /// </summary>
    public interface IDateCalculator
    {
        /// <summary>
        /// Adds a specified number of days to a date.
        /// </summary>
        /// <param name="date">The starting date.</param>
        /// <param name="days">The number of days to add (can be negative).</param>
        /// <returns>The resulting date after adding the specified days.</returns>
        DateTime AddDays(DateTime date, int days);

        /// <summary>
        /// Gets the Italian month abbreviation for a given month number.
        /// </summary>
        /// <param name="month">The month number (1-12).</param>
        /// <returns>The Italian month abbreviation (gen, feb, mar, apr, mag, giu, lug, ago, set, ott, nov, dic).</returns>
        /// <exception cref="ArgumentOutOfRangeException">Thrown when month is not between 1 and 12.</exception>
        string GetItalianMonthAbbreviation(int month);

        /// <summary>
        /// Formats a date in Italian format with day number and month abbreviation.
        /// </summary>
        /// <param name="date">The date to format.</param>
        /// <returns>A string in the format "DD mmm" (e.g., "26 gen").</returns>
        string FormatItalianDate(DateTime date);

        /// <summary>
        /// Parses an Italian date string into a DateTime object.
        /// </summary>
        /// <param name="dateText">The date text in format "DD mmm" (e.g., "26 gen").</param>
        /// <param name="year">The year to use for the parsed date.</param>
        /// <returns>The parsed DateTime object.</returns>
        /// <exception cref="FormatException">Thrown when the date text cannot be parsed.</exception>
        DateTime ParseItalianDate(string dateText, int year);
    }
}
