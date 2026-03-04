using System.Collections.Generic;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for CSV parsing service.
    /// Responsible for reading and parsing CSV files into ServiceAppointment objects.
    /// Validates: Requirements 2.1, 2.2, 2.4
    /// </summary>
    public interface ICSVParser
    {
        /// <summary>
        /// Parses a CSV file and returns a list of ServiceAppointment objects.
        /// Supports UTF-8 encoding for Italian character preservation.
        /// </summary>
        /// <param name="filePath">The path to the CSV file to parse</param>
        /// <returns>A list of ServiceAppointment objects parsed from the CSV file</returns>
        /// <exception cref="System.IO.FileNotFoundException">Thrown when the CSV file is not found</exception>
        /// <exception cref="System.IO.IOException">Thrown when the CSV file cannot be read</exception>
        /// <exception cref="CsvHelper.CsvHelperException">Thrown when the CSV file is malformed</exception>
        List<ServiceAppointment> ParseCSV(string filePath);

        /// <summary>
        /// Validates that a CSV file contains all required columns.
        /// </summary>
        /// <param name="filePath">The path to the CSV file to validate</param>
        /// <returns>True if the CSV file has all required columns, false otherwise</returns>
        bool ValidateCSVStructure(string filePath);

        /// <summary>
        /// Validates that a CSV file contains all required columns and returns detailed error information.
        /// </summary>
        /// <param name="filePath">The path to the CSV file to validate</param>
        /// <param name="missingColumns">Output parameter containing the list of missing column names</param>
        /// <returns>True if the CSV file has all required columns, false otherwise</returns>
        bool ValidateCSVStructure(string filePath, out List<string> missingColumns);
    }
}
