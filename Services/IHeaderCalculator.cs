using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for header calculation operations.
    /// Handles parsing existing headers and generating new week headers.
    /// </summary>
    public interface IHeaderCalculator
    {
        /// <summary>
        /// Parses a header string to extract date range, week number, and referente information.
        /// Expected format: "DD mmm DD mmm Settimana Nreferente settimana = [text]"
        /// Example: "26 gen 01 feb Settimana 5referente settimana = Inserire nome e numero di telefono del referente"
        /// </summary>
        /// <param name="headerText">The header text to parse.</param>
        /// <returns>A HeaderInfo object containing the parsed information.</returns>
        /// <exception cref="FormatException">Thrown when the header text cannot be parsed.</exception>
        HeaderInfo ParseHeader(string headerText);

        /// <summary>
        /// Generates the next week's header based on the previous week's header.
        /// Adds 7 days to both Monday and Sunday dates, increments week number by 1,
        /// and resets the referente text to the default value.
        /// </summary>
        /// <param name="previousHeader">The previous week's header text.</param>
        /// <returns>The formatted header text for the next week.</returns>
        /// <exception cref="FormatException">Thrown when the previous header cannot be parsed.</exception>
        string GenerateNextWeekHeader(string previousHeader);
    }
}
