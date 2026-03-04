using System.Collections.Generic;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for performing VLOOKUP-style operations against reference sheets.
    /// </summary>
    public interface ILookupService
    {
        /// <summary>
        /// Performs a lookup in the assistiti reference sheet.
        /// </summary>
        /// <param name="assistitoName">The assistito name to look up</param>
        /// <param name="columnName">The column name to retrieve</param>
        /// <returns>The value from the reference sheet, or empty string if not found</returns>
        string LookupInAssistiti(string assistitoName, string columnName);

        /// <summary>
        /// Performs a lookup in the fissi reference sheet.
        /// </summary>
        /// <param name="lookupKey">The lookup key</param>
        /// <param name="columnName">The column name to retrieve</param>
        /// <returns>The value from the reference sheet, or empty string if not found</returns>
        string LookupInFissi(string lookupKey, string columnName);

        /// <summary>
        /// Loads and caches reference sheet data for O(1) lookups.
        /// </summary>
        /// <param name="assistitiSheet">The assistiti reference sheet</param>
        /// <param name="fissiSheet">The fissi reference sheet</param>
        void LoadReferenceSheets(Sheet assistitiSheet, Sheet fissiSheet);
    }

    /// <summary>
    /// Implementation of ILookupService that performs VLOOKUP-style operations
    /// against reference sheets with O(1) lookup performance.
    /// </summary>
    public class LookupService : ILookupService
    {
        // Cache structure: Dictionary<lookupKey, Dictionary<columnName, value>>
        private Dictionary<string, Dictionary<string, string>> _assistitiData;
        private Dictionary<string, Dictionary<string, string>> _fissiData;

        public LookupService()
        {
            _assistitiData = new Dictionary<string, Dictionary<string, string>>();
            _fissiData = new Dictionary<string, Dictionary<string, string>>();
        }

        /// <summary>
        /// Loads reference sheets once at the start of processing and caches data in memory.
        /// </summary>
        public void LoadReferenceSheets(Sheet assistitiSheet, Sheet fissiSheet)
        {
            _assistitiData = LoadSheetData(assistitiSheet);
            _fissiData = LoadSheetData(fissiSheet);
        }

        /// <summary>
        /// Performs a lookup in the assistiti reference sheet.
        /// Returns empty string for missing keys (not null).
        /// </summary>
        public string LookupInAssistiti(string assistitoName, string columnName)
        {
            return PerformLookup(_assistitiData, assistitoName, columnName);
        }

        /// <summary>
        /// Performs a lookup in the fissi reference sheet.
        /// Returns empty string for missing keys (not null).
        /// </summary>
        public string LookupInFissi(string lookupKey, string columnName)
        {
            return PerformLookup(_fissiData, lookupKey, columnName);
        }

        /// <summary>
        /// Loads data from a sheet into a Dictionary structure for O(1) lookups.
        /// First row is assumed to be headers.
        /// First column is assumed to be the lookup key.
        /// </summary>
        private Dictionary<string, Dictionary<string, string>> LoadSheetData(Sheet sheet)
        {
            var data = new Dictionary<string, Dictionary<string, string>>();
            
            if (sheet?.Worksheet == null)
                return data;

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            
            if (dimension == null)
                return data;

            // Read header row (row 1)
            var headers = new List<string>();
            for (int col = 1; col <= dimension.End.Column; col++)
            {
                var headerValue = worksheet.Cells[1, col].Value?.ToString() ?? "";
                headers.Add(headerValue);
            }

            // Read data rows (starting from row 2)
            for (int row = 2; row <= dimension.End.Row; row++)
            {
                // First column is the lookup key
                var lookupKey = worksheet.Cells[row, 1].Value?.ToString() ?? "";
                
                if (string.IsNullOrWhiteSpace(lookupKey))
                    continue;

                var rowData = new Dictionary<string, string>();
                
                // Read all columns for this row
                for (int col = 1; col <= dimension.End.Column; col++)
                {
                    var columnName = headers[col - 1];
                    var cellValue = worksheet.Cells[row, col].Value?.ToString() ?? "";
                    rowData[columnName] = cellValue;
                }

                // Store using first match (if multiple matches exist, use the first)
                if (!data.ContainsKey(lookupKey))
                {
                    data[lookupKey] = rowData;
                }
            }

            return data;
        }

        /// <summary>
        /// Performs a lookup operation on cached data.
        /// Returns empty string if key or column not found.
        /// </summary>
        private string PerformLookup(
            Dictionary<string, Dictionary<string, string>> cache,
            string lookupKey,
            string columnName)
        {
            if (string.IsNullOrWhiteSpace(lookupKey))
                return "";

            if (!cache.ContainsKey(lookupKey))
                return "";

            var rowData = cache[lookupKey];
            
            if (!rowData.ContainsKey(columnName))
                return "";

            return rowData[columnName];
        }
    }
}
