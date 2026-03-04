using System.Collections.Generic;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for managing the new column layout and positioning.
    /// </summary>
    public interface IColumnStructureManager
    {
        /// <summary>
        /// Gets the column headers in the new 15-column structure.
        /// </summary>
        /// <returns>List of column header names in order</returns>
        List<string> GetColumnHeaders();

        /// <summary>
        /// Gets the zero-based index of a column by name.
        /// </summary>
        /// <param name="columnName">The column name to look up</param>
        /// <returns>The zero-based column index, or -1 if not found</returns>
        int GetColumnIndex(string columnName);

        /// <summary>
        /// Maps an old column name to the new column name.
        /// </summary>
        /// <param name="oldColumnName">The old column name</param>
        /// <returns>The new column name, or the original name if no mapping exists</returns>
        string GetNewColumnName(string oldColumnName);
    }

    /// <summary>
    /// Implementation of IColumnStructureManager that defines the new 15-column structure
    /// for the enhanced Excel output.
    /// </summary>
    public class ColumnStructureManager : IColumnStructureManager
    {
        private readonly List<string> _columnHeaders;
        private readonly Dictionary<string, string> _columnNameMapping;

        public ColumnStructureManager()
        {
            // Define the 12-column structure
            _columnHeaders = new List<string>
            {
                "Data",                 // 1
                "Partenza",             // 2 (renamed from "Ora Inizio Servizio")
                "Assistito",            // 3
                "Indirizzo",            // 4 (from assistiti lookup, positioned after Assistito)
                "Destinazione",         // 5
                "Note",                 // 6 (from assistiti lookup)
                "Auto",                 // 7
                "Volontario",           // 8
                "Arrivo",               // 9
                "Avv",                  // 10 (new, from fissi lookup)
                "Indirizzo Gasnet",     // 11
                "Note Gasnet"           // 12 (from CSV)
            };

            // Define old-to-new column name mappings
            _columnNameMapping = new Dictionary<string, string>
            {
                { "Ora Inizio Servizio", "Partenza" }
            };
        }

        /// <summary>
        /// Returns the 12-column header structure.
        /// </summary>
        public List<string> GetColumnHeaders()
        {
            return new List<string>(_columnHeaders);
        }

        /// <summary>
        /// Gets the zero-based index of a column by name.
        /// Returns -1 if the column is not found.
        /// </summary>
        public int GetColumnIndex(string columnName)
        {
            if (string.IsNullOrWhiteSpace(columnName))
                return -1;

            return _columnHeaders.IndexOf(columnName);
        }

        /// <summary>
        /// Maps an old column name to the new column name.
        /// Returns the original name if no mapping exists.
        /// </summary>
        public string GetNewColumnName(string oldColumnName)
        {
            if (string.IsNullOrWhiteSpace(oldColumnName))
                return oldColumnName;

            if (_columnNameMapping.ContainsKey(oldColumnName))
                return _columnNameMapping[oldColumnName];

            return oldColumnName;
        }
    }
}
