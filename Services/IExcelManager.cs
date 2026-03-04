using System.Collections.Generic;
using AuserExcelTransformer.Models;
using ExcelWorkbookModel = AuserExcelTransformer.Models.ExcelWorkbook;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for Excel workbook management operations.
    /// Handles reading, writing, and manipulating Excel files for the Auser Excel Transformer.
    /// </summary>
    public interface IExcelManager
    {
        /// <summary>
        /// Opens an Excel workbook from the specified file path.
        /// </summary>
        /// <param name="filePath">Path to the Excel file (.xlsx)</param>
        /// <returns>ExcelWorkbook wrapper object</returns>
        /// <exception cref="System.IO.FileNotFoundException">If the file does not exist</exception>
        /// <exception cref="System.InvalidOperationException">If the file cannot be opened</exception>
        ExcelWorkbookModel OpenWorkbook(string filePath);

        /// <summary>
        /// Gets all sheet names from the workbook.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <returns>List of sheet names</returns>
        List<string> GetSheetNames(ExcelWorkbookModel workbook);

        /// <summary>
        /// Determines the next sequential sheet number by finding the highest numbered sheet and incrementing by one.
        /// </summary>
        /// <param name="sheetNames">List of existing sheet names</param>
        /// <returns>Next sheet number (max + 1, or 1 if no numbered sheets exist)</returns>
        int GetNextSheetNumber(List<string> sheetNames);

        /// <summary>
        /// Locates and returns the "fissi" sheet from the workbook.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <returns>Sheet wrapper for the fissi sheet</returns>
        /// <exception cref="System.InvalidOperationException">If the fissi sheet is not found</exception>
        Sheet GetFissiSheet(ExcelWorkbookModel workbook);

        /// <summary>
        /// Creates a new sheet in the workbook with the specified number.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <param name="sheetNumber">The sheet number for naming</param>
        /// <returns>The newly created sheet</returns>
        Sheet CreateNewSheet(ExcelWorkbookModel workbook, int sheetNumber);

        /// <summary>
        /// Writes the header row (row 1) with dates, week number, and referente.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="mondayDate">The Monday date for the week</param>
        void WriteHeaderRow(Sheet sheet, System.DateTime mondayDate);

        /// <summary>
        /// Writes the column headers in row 2.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        void WriteColumnHeaders(Sheet sheet);

        /// <summary>
        /// Writes transformed data rows to the sheet starting at the specified row.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rows">List of transformed rows to write</param>
        /// <param name="startRow">Starting row number (1-based)</param>
        void WriteDataRows(Sheet sheet, List<TransformedRow> rows, int startRow);

        /// <summary>
        /// Appends data from the fissi sheet to the target sheet, preserving formatting.
        /// </summary>
        /// <param name="targetSheet">The sheet to append data to</param>
        /// <param name="fissiSheet">The fissi sheet to copy from</param>
        /// <param name="startRow">Starting row number in target sheet (1-based)</param>
        void AppendFissiData(Sheet targetSheet, Sheet fissiSheet, int startRow);

        /// <summary>
        /// Applies yellow highlighting to the specified rows.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rowNumbers">List of row numbers to highlight (1-based)</param>
        void ApplyYellowHighlight(Sheet sheet, List<int> rowNumbers);

        /// <summary>
        /// Enables AutoFilter for the data range starting from row 2 (column headers).
        /// This allows users to filter and sort data by any column.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        void EnableAutoFilter(Sheet sheet);

        /// <summary>
        /// Saves the workbook to the specified file path.
        /// </summary>
        /// <param name="workbook">The Excel workbook to save</param>
        /// <param name="filePath">Path where the file should be saved</param>
        void SaveWorkbook(ExcelWorkbookModel workbook, string filePath);

        /// <summary>
        /// Gets a sheet by its name.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <param name="sheetName">The name of the sheet to retrieve</param>
        /// <returns>The sheet with the specified name, or null if not found</returns>
        Sheet GetSheetByName(ExcelWorkbookModel workbook, string sheetName);

        /// <summary>
        /// Reads the header text from the first row of a sheet.
        /// </summary>
        /// <param name="sheet">The sheet to read from</param>
        /// <returns>The header text from the first cell of the first row</returns>
        string ReadHeader(Sheet sheet);

        /// <summary>
        /// Writes the enhanced column headers in row 2 using the new 15-column structure.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        void WriteColumnHeadersEnhanced(Sheet sheet);

        /// <summary>
        /// Writes enhanced transformed data rows to the sheet starting at the specified row.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rows">List of enhanced transformed rows to write</param>
        /// <param name="startRow">Starting row number (1-based)</param>
        void WriteDataRowsEnhanced(Sheet sheet, List<EnhancedTransformedRow> rows, int startRow);

        /// <summary>
        /// Sorts data rows by Data column (primary, ascending) and Partenza column (secondary, ascending).
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="startRow">First data row (1-based)</param>
        /// <param name="endRow">Last data row (1-based)</param>
        void SortDataRows(Sheet sheet, int startRow, int endRow);

        /// <summary>
        /// Applies bold formatting to column headers.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="headerRow">The row number of the header (1-based)</param>
        void ApplyBoldToHeaders(Sheet sheet, int headerRow);

        /// <summary>
        /// Applies thick borders to the last row of each date group.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="startRow">First data row (1-based)</param>
        /// <param name="endRow">Last data row (1-based)</param>
        void ApplyThickBordersToDateGroups(Sheet sheet, int startRow, int endRow);
    }
}
