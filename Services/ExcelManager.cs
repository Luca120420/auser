using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using ExcelWorkbookModel = AuserExcelTransformer.Models.ExcelWorkbook;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Implementation of IExcelManager for managing Excel workbook operations using EPPlus.
    /// Handles opening workbooks, reading sheets, creating new sheets, and writing data.
    /// </summary>
    public class ExcelManager : IExcelManager
    {
        /// <summary>
        /// Opens an Excel workbook from the specified file path.
        /// </summary>
        /// <param name="filePath">Path to the Excel file (.xlsx)</param>
        /// <returns>ExcelWorkbook wrapper object</returns>
        /// <exception cref="FileNotFoundException">If the file does not exist</exception>
        /// <exception cref="InvalidOperationException">If the file cannot be opened</exception>
        public ExcelWorkbookModel OpenWorkbook(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Il file Excel non è stato trovato: {filePath}");
            }

            try
            {
                // Set EPPlus license context (required for EPPlus 5.0+)
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                var fileInfo = new FileInfo(filePath);
                var package = new ExcelPackage(fileInfo);
                
                return new ExcelWorkbookModel(package);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Impossibile aprire il file Excel: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Gets all sheet names from the workbook.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <returns>List of sheet names</returns>
        public List<string> GetSheetNames(ExcelWorkbookModel workbook)
        {
            if (workbook == null || workbook.Package == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var sheetNames = new List<string>();
            foreach (var worksheet in workbook.Package.Workbook.Worksheets)
            {
                sheetNames.Add(worksheet.Name);
            }

            return sheetNames;
        }

        /// <summary>
        /// Determines the next sequential sheet number by finding the highest numbered sheet and incrementing by one.
        /// Looks for sheets with numeric names (e.g., "1", "2", "3") and returns max + 1.
        /// Only considers sheet numbers in the valid range (1-53) to avoid confusion with year sheets.
        /// </summary>
        /// <param name="sheetNames">List of existing sheet names</param>
        /// <returns>Next sheet number (max + 1, or 1 if no numbered sheets exist)</returns>
        public int GetNextSheetNumber(List<string> sheetNames)
        {
            if (sheetNames == null || sheetNames.Count == 0)
            {
                return 1;
            }

            int maxNumber = 0;

            foreach (var sheetName in sheetNames)
            {
                // Try to parse the sheet name as an integer
                if (int.TryParse(sheetName.Trim(), out int number))
                {
                    // Only consider numbers in valid week range (1-53)
                    // This avoids treating year sheets (2025, 2026) as weekly sheets
                    if (number >= 1 && number <= 53 && number > maxNumber)
                    {
                        maxNumber = number;
                    }
                }
            }

            return maxNumber + 1;
        }

        /// <summary>
        /// Locates and returns the "fissi" sheet from the workbook.
        /// Performs case-insensitive search for the sheet named "fissi".
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <returns>Sheet wrapper for the fissi sheet</returns>
        /// <exception cref="InvalidOperationException">If the fissi sheet is not found</exception>
        public Sheet GetFissiSheet(ExcelWorkbookModel workbook)
        {
            if (workbook == null || workbook.Package == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            // Search for "fissi" sheet (case-insensitive)
            var fissiWorksheet = workbook.Package.Workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.Equals("fissi", StringComparison.OrdinalIgnoreCase));

            if (fissiWorksheet == null)
            {
                throw new InvalidOperationException("Il foglio 'fissi' non è stato trovato nel file Excel.");
            }

            return new Sheet(fissiWorksheet);
        }

        /// <summary>
        /// Creates a new sheet in the workbook with the specified number.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <param name="sheetNumber">The sheet number for naming</param>
        /// <returns>The newly created sheet</returns>
        public Sheet CreateNewSheet(ExcelWorkbookModel workbook, int sheetNumber)
        {
            if (workbook == null || workbook.Package == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var worksheet = workbook.Package.Workbook.Worksheets.Add(sheetNumber.ToString());
            return new Sheet(worksheet);
        }

        /// <summary>
        /// Writes the header row (row 1) with dates, week number, and referente.
        /// A1: Date value (Monday)
        /// B1: Formula =A1+6 (Sunday)
        /// C1: Formula =CONCATENATE("Settimana ",WEEKNUM(A1))
        /// D1: "referente settimana = Inserire nome e numero di telefono del referente"
        /// All cells formatted as Calibri 16 Bold
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="mondayDate">The Monday date for the week</param>
        public void WriteHeaderRow(Sheet sheet, DateTime mondayDate)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var ws = sheet.Worksheet;

            // A1: Monday date as Excel date value
            ws.Cells[1, 1].Value = mondayDate;
            ws.Cells[1, 1].Style.Numberformat.Format = "dd mmm";
            ws.Cells[1, 1].Style.Font.Name = "Calibri";
            ws.Cells[1, 1].Style.Font.Size = 16;
            ws.Cells[1, 1].Style.Font.Bold = true;

            // B1: Formula =A1+6 (Sunday)
            ws.Cells[1, 2].Formula = "A1+6";
            ws.Cells[1, 2].Style.Numberformat.Format = "dd mmm";
            ws.Cells[1, 2].Style.Font.Name = "Calibri";
            ws.Cells[1, 2].Style.Font.Size = 16;
            ws.Cells[1, 2].Style.Font.Bold = true;

            // C1: Formula =CONCATENATE("Settimana ",WEEKNUM(A1))
            ws.Cells[1, 3].Formula = "CONCATENATE(\"Settimana \",WEEKNUM(A1))";
            ws.Cells[1, 3].Style.Font.Name = "Calibri";
            ws.Cells[1, 3].Style.Font.Size = 16;
            ws.Cells[1, 3].Style.Font.Bold = true;

            // D1: Referente text
            ws.Cells[1, 4].Value = "referente settimana = Inserire nome e numero di telefono del referente";
            ws.Cells[1, 4].Style.Font.Name = "Calibri";
            ws.Cells[1, 4].Style.Font.Size = 16;
            ws.Cells[1, 4].Style.Font.Bold = true;
        }

        /// <summary>
        /// Writes the column headers in row 2.
        /// New structure: Data, Ora Inizio Servizio, Assistito, Destinazione, Note, Auto, Volontario, Arrivo, [empty], Indirizzo Partenza, Comune Partenza, [2 empty], Indirizzo Gasnet
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        public void WriteColumnHeaders(Sheet sheet)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var ws = sheet.Worksheet;
            string[] headers = { 
                "Data", 
                "Ora Inizio Servizio", 
                "Assistito", 
                "Destinazione", 
                "Note", 
                "Auto", 
                "Volontario", 
                "Arrivo", 
                "", // Empty column 1
                "Indirizzo Partenza", 
                "Comune Partenza", 
                "", // Empty column 2
                "", // Empty column 3
                "Indirizzo Gasnet" 
            };

            for (int i = 0; i < headers.Length; i++)
            {
                ws.Cells[2, i + 1].Value = headers[i];
            }
        }

        /// <summary>
        /// Writes transformed data rows to the sheet starting at the specified row.
        /// New structure: Data, Ora Inizio Servizio, Assistito, Destinazione, Note, Auto, Volontario, Arrivo, [empty], Indirizzo Partenza, Comune Partenza, [2 empty], Indirizzo Gasnet
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rows">List of transformed rows to write</param>
        /// <param name="startRow">Starting row number (1-based)</param>
        public void WriteDataRows(Sheet sheet, List<TransformedRow> rows, int startRow)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (rows == null)
            {
                return;
            }

            int currentRow = startRow;
            foreach (var row in rows)
            {
                int col = 1;
                
                // Column 1: Data - parse and format as date with "ddd dd mmm" format (e.g., "gio 12 feb")
                var dataCell = sheet.Worksheet.Cells[currentRow, col++];
                DateTime dataDate;
                if (DateTime.TryParseExact(row.DataServizio, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" }, 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out dataDate))
                {
                    dataCell.Value = dataDate;
                    dataCell.Style.Numberformat.Format = "ddd dd mmm";
                }
                else if (DateTime.TryParse(row.DataServizio, out dataDate))
                {
                    dataCell.Value = dataDate;
                    dataCell.Style.Numberformat.Format = "ddd dd mmm";
                }
                else
                {
                    dataCell.Value = row.DataServizio;
                }
                
                // Column 2: Ora Inizio Servizio (empty for CSV)
                var partenzaCell = sheet.Worksheet.Cells[currentRow, col++];
                partenzaCell.Value = row.OraInizioServizio;
                partenzaCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                
                // Column 3: Assistito
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Assistito;
                
                // Column 4: Destinazione
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Destinazione;
                
                // Column 5: Note
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Note;
                
                // Column 6: Auto
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Auto;
                
                // Column 7: Volontario
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Volontario;
                
                // Column 8: Arrivo
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Arrivo;
                
                // Column 9: Empty column
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Empty1;
                
                // Column 10: Indirizzo Partenza
                sheet.Worksheet.Cells[currentRow, col++].Value = row.IndirizzoPartenza;
                
                // Column 11: Comune Partenza
                sheet.Worksheet.Cells[currentRow, col++].Value = row.ComunePartenza;
                
                // Columns 12-13: Empty columns
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Empty2;
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Empty3;
                
                // Column 14: Indirizzo Gasnet
                sheet.Worksheet.Cells[currentRow, col++].Value = row.IndirizzoGasnet;

                currentRow++;
            }
        }

        /// <summary>
        /// Appends data from the fissi sheet to the target sheet, preserving formatting but not background colors.
        /// Skips header rows from the fissi sheet (detects if row 2 contains "Data" header).
        /// Maps fissi sheet columns (9 columns old structure) to new 14-column structure:
        /// Fissi: Data, Partenza, Assistito, Indirizzo, Destinazione, Note, Auto, Volontario, Arrivo
        /// New:   Data, [empty], Assistito, Indirizzo, Destinazione, [3 empty], Arrivo, [empty], Partenza, [empty], Note, [empty]
        /// </summary>
        /// <param name="targetSheet">The sheet to append data to</param>
        /// <param name="fissiSheet">The fissi sheet to copy from</param>
        /// <param name="startRow">Starting row number in target sheet (1-based)</param>
        public void AppendFissiData(Sheet targetSheet, Sheet fissiSheet, int startRow)
        {
            if (targetSheet == null || targetSheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(targetSheet));
            }

            if (fissiSheet == null || fissiSheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(fissiSheet));
            }

            var fissiWorksheet = fissiSheet.Worksheet;
            var targetWorksheet = targetSheet.Worksheet;

            // Get the dimensions of the fissi sheet
            var fissiDimension = fissiWorksheet.Dimension;
            if (fissiDimension == null)
            {
                return; // Empty sheet
            }

            // Detect where data starts in fissi sheet
            // Check if row 2 contains "Data" in cell A2 (column header)
            int fissiDataStartRow = 3; // Default: skip rows 1 and 2
            
            var row2Cell = fissiWorksheet.Cells[2, 1];
            string row2Text = row2Cell.Text?.Trim() ?? string.Empty;
            
            // If row 2 contains "Data", it's a column header row, so skip it and start from row 3
            if (row2Text.Equals("Data", StringComparison.OrdinalIgnoreCase))
            {
                fissiDataStartRow = 3;
            }
            else
            {
                // If row 2 doesn't contain "Data", check row 1
                var row1Cell = fissiWorksheet.Cells[1, 1];
                string row1Text = row1Cell.Text?.Trim() ?? string.Empty;
                
                if (row1Text.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    // Row 1 is column headers, start from row 2
                    fissiDataStartRow = 2;
                }
                else
                {
                    // Neither row 1 nor row 2 contains "Data", assume row 2 is data
                    fissiDataStartRow = 2;
                }
            }

            // Column mapping from fissi (9 cols) to new structure (14 cols):
            // Fissi Col 1 (Data) → Target Col 1 (Data)
            // Fissi Col 2 (Partenza) → Target Col 2 (Ora Inizio Servizio)
            // Fissi Col 3 (Assistito) → Target Col 3 (Assistito)
            // Fissi Col 4 (Indirizzo) → Target Col 4 (Indirizzo)
            // Fissi Col 5 (Destinazione) → Target Col 5 (Destinazione)
            // Fissi Col 6 (Note) → Target Col 6 (Note)
            // Fissi Col 7 (Auto) → Target Col 7 (Auto)
            // Fissi Col 8 (Volontario) → Target Col 8 (Volontario)
            // Fissi Col 9 (Arrivo) → Target Col 9 (Arrivo)
            // Target Cols 10-14 remain empty for other data sources
            
            int[] columnMapping = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9 };

            // Copy data rows only (skip header rows)
            int targetRow = startRow;
            for (int sourceRow = fissiDataStartRow; sourceRow <= fissiDimension.End.Row; sourceRow++)
            {
                // Copy mapped columns from fissi sheet
                for (int fissiCol = 1; fissiCol <= Math.Min(9, fissiDimension.End.Column); fissiCol++)
                {
                    int targetCol = columnMapping[fissiCol - 1];
                    
                    var sourceCell = fissiWorksheet.Cells[sourceRow, fissiCol];
                    var targetCell = targetWorksheet.Cells[targetRow, targetCol];

                    // Special handling for time columns (2 = Partenza, 9 = Arrivo)
                    if (fissiCol == 2 || fissiCol == 9)
                    {
                        // Handle different time formats in source data
                        var sourceValue = sourceCell.Value;
                        
                        if (sourceValue is double || sourceValue is decimal)
                        {
                            // Already a numeric time value (fraction of a day)
                            targetCell.Value = sourceValue;
                        }
                        else if (sourceValue is DateTime dt)
                        {
                            // DateTime object - convert to Excel time fraction
                            targetCell.Value = dt.TimeOfDay.TotalDays;
                        }
                        else if (sourceValue is string strValue && !string.IsNullOrWhiteSpace(strValue))
                        {
                            // Text value - try to parse as time
                            // Handle formats like "8.30.00", "8:30:00", "8:30", "08:30", etc.
                            string normalizedTime = strValue.Trim()
                                .Replace('.', ':')  // Convert dots to colons
                                .Replace(',', ':'); // Convert commas to colons
                            
                            if (TimeSpan.TryParse(normalizedTime, out TimeSpan timeSpan))
                            {
                                // Convert TimeSpan to Excel time fraction (fraction of a day)
                                targetCell.Value = timeSpan.TotalDays;
                            }
                            else
                            {
                                // Can't parse - copy as-is
                                targetCell.Value = sourceValue;
                            }
                        }
                        else
                        {
                            // Null or other type - copy as-is
                            targetCell.Value = sourceValue;
                        }
                    }
                    else
                    {
                        // Non-time columns - copy value as-is
                        targetCell.Value = sourceCell.Value;
                    }

                    // Copy formatting (but NEVER copy background color to avoid yellow highlighting)
                    if (sourceCell.Style != null)
                    {
                        // Copy font properties
                        targetCell.Style.Font.Bold = sourceCell.Style.Font.Bold;
                        targetCell.Style.Font.Italic = sourceCell.Style.Font.Italic;
                        targetCell.Style.Font.Size = sourceCell.Style.Font.Size;
                        targetCell.Style.Font.Name = sourceCell.Style.Font.Name;
                        
                        // Copy font color if set
                        if (!string.IsNullOrEmpty(sourceCell.Style.Font.Color.Rgb))
                        {
                            var fontColorHex = sourceCell.Style.Font.Color.Rgb;
                            targetCell.Style.Font.Color.SetColor(System.Drawing.ColorTranslator.FromHtml("#" + fontColorHex.Substring(2)));
                        }
                        
                        // EXPLICITLY SKIP background color to avoid yellow highlighting
                        // Do NOT copy: sourceCell.Style.Fill.PatternType and BackgroundColor
                        
                        // Copy borders
                        targetCell.Style.Border.Top.Style = sourceCell.Style.Border.Top.Style;
                        targetCell.Style.Border.Bottom.Style = sourceCell.Style.Border.Bottom.Style;
                        targetCell.Style.Border.Left.Style = sourceCell.Style.Border.Left.Style;
                        targetCell.Style.Border.Right.Style = sourceCell.Style.Border.Right.Style;
                        
                        // Copy number format (for dates and other formatted columns) - exclude time columns 2 and 9
                        if (!string.IsNullOrEmpty(sourceCell.Style.Numberformat.Format) && fissiCol != 2 && fissiCol != 9)
                        {
                            targetCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
                        }
                        
                        // Always apply time format to columns 2 (Partenza) and 9 (Arrivo)
                        if (fissiCol == 2 || fissiCol == 9)
                        {
                            targetCell.Style.Numberformat.Format = "h:mm";
                        }
                    }
                }
                targetRow++;
            }
        }

        /// <summary>
        /// Applies yellow highlighting to the specified rows.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rowNumbers">List of row numbers to highlight (1-based)</param>
        public void ApplyYellowHighlight(Sheet sheet, List<int> rowNumbers)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (rowNumbers == null || rowNumbers.Count == 0)
            {
                return;
            }

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            if (dimension == null)
            {
                return;
            }

            foreach (var rowNumber in rowNumbers)
            {
                // Apply yellow background to all cells in the row
                for (int col = 1; col <= dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[rowNumber, col];
                    cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                    cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                }
            }
        }

        /// <summary>
        /// Enables AutoFilter for the data range starting from row 2 (column headers).
        /// This allows users to filter and sort data by any column.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        public void EnableAutoFilter(Sheet sheet)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            
            if (dimension == null || dimension.End.Row < 2)
            {
                return; // Not enough data to apply filter
            }

            // Apply AutoFilter starting from row 2 (column headers) to the last row
            worksheet.Cells[2, 1, dimension.End.Row, dimension.End.Column].AutoFilter = true;
        }

        /// <summary>
        /// Writes the enhanced column headers in row 2 using the new 14-column structure.
        /// Structure: Data, Partenza, Assistito, Indirizzo, Destinazione, Note, Auto, Volontario, 
        /// Arrivo, Avv, [empty], Indirizzo Gasnet, Note Gasnet, [empty]
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        public void WriteColumnHeadersEnhanced(Sheet sheet)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var columnStructureManager = new ColumnStructureManager();
            var headers = columnStructureManager.GetColumnHeaders();
            var ws = sheet.Worksheet;

            for (int i = 0; i < headers.Count; i++)
            {
                ws.Cells[2, i + 1].Value = headers[i];
            }
        }

        /// <summary>
        /// Writes enhanced transformed data rows to the sheet starting at the specified row.
        /// Uses the new 14-column structure with all enhanced fields.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="rows">List of enhanced transformed rows to write</param>
        /// <param name="startRow">Starting row number (1-based)</param>
        public void WriteDataRowsEnhanced(Sheet sheet, List<EnhancedTransformedRow> rows, int startRow)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (rows == null)
            {
                return;
            }

            int currentRow = startRow;
            foreach (var row in rows)
            {
                int col = 1;
                
                // Column 1: Data - parse and format as date with "ddd dd mmm" format
                var dataCell = sheet.Worksheet.Cells[currentRow, col++];
                DateTime dataDate;
                if (DateTime.TryParseExact(row.Data, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" }, 
                    System.Globalization.CultureInfo.InvariantCulture, 
                    System.Globalization.DateTimeStyles.None, out dataDate))
                {
                    dataCell.Value = dataDate;
                    dataCell.Style.Numberformat.Format = "ddd dd mmm";
                }
                else if (DateTime.TryParse(row.Data, out dataDate))
                {
                    dataCell.Value = dataDate;
                    dataCell.Style.Numberformat.Format = "ddd dd mmm";
                }
                else
                {
                    dataCell.Value = row.Data;
                }
                
                // Column 2: Partenza
                var partenzaCell = sheet.Worksheet.Cells[currentRow, col++];
                partenzaCell.Value = row.Partenza;
                partenzaCell.Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                
                // Column 3: Assistito
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Assistito;
                
                // Column 4: Indirizzo (from assistiti lookup)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Indirizzo;
                
                // Column 5: Destinazione
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Destinazione;
                
                // Column 6: Note (from assistiti lookup)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Note;
                
                // Column 7: Auto
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Auto;
                
                // Column 8: Volontario
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Volontario;
                
                // Column 9: Arrivo
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Arrivo;
                
                // Column 10: Avv (from fissi lookup)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Avv;
                
                // Column 11: Empty1
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Empty1;
                
                // Column 12: Indirizzo Gasnet
                sheet.Worksheet.Cells[currentRow, col++].Value = row.IndirizzoGasnet;
                
                // Column 13: Note Gasnet (from CSV)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.NoteGasnet;
                
                // Column 14: Empty2
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Empty2;

                currentRow++;
            }
        }

        /// <summary>
        /// Sorts data rows by Data column (primary, ascending) and Partenza column (secondary, ascending).
        /// Handles invalid date/time formats gracefully with Italian error messages.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="startRow">First data row (1-based)</param>
        /// <param name="endRow">Last data row (1-based)</param>
        public void SortDataRows(Sheet sheet, int startRow, int endRow)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (startRow < 1 || endRow < startRow)
            {
                return;
            }

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            
            if (dimension == null || startRow > endRow)
            {
                return;
            }

            try
            {
                // Read all rows into memory for sorting
                var rows = new List<(int rowNum, DateTime date, TimeSpan time, object[] values)>();
                
                for (int row = startRow; row <= endRow; row++)
                {
                    // Get date from column 1
                    var dateCell = worksheet.Cells[row, 1];
                    DateTime date;
                    
                    if (dateCell.Value is DateTime)
                    {
                        date = (DateTime)dateCell.Value;
                    }
                    else if (dateCell.Value is double)
                    {
                        date = DateTime.FromOADate((double)dateCell.Value);
                    }
                    else if (DateTime.TryParse(dateCell.Value?.ToString(), out DateTime parsedDate))
                    {
                        date = parsedDate;
                    }
                    else
                    {
                        continue; // Skip rows with invalid dates
                    }
                    
                    // Get time from column 2
                    var timeCell = worksheet.Cells[row, 2];
                    TimeSpan time = TimeSpan.Zero;
                    
                    if (timeCell.Value != null)
                    {
                        var timeStr = timeCell.Value.ToString();
                        if (!TimeSpan.TryParse(timeStr, out time))
                        {
                            // Try parsing HH:mm format
                            if (timeStr.Contains(":"))
                            {
                                var parts = timeStr.Split(':');
                                if (parts.Length >= 2 &&
                                    int.TryParse(parts[0], out int hours) &&
                                    int.TryParse(parts[1], out int minutes))
                                {
                                    time = new TimeSpan(hours, minutes, 0);
                                }
                            }
                        }
                    }
                    
                    // Read all cell values for this row
                    var values = new object[dimension.End.Column];
                    for (int col = 1; col <= dimension.End.Column; col++)
                    {
                        values[col - 1] = worksheet.Cells[row, col].Value;
                    }
                    
                    rows.Add((row, date, time, values));
                }
                
                // Sort by date (primary) then time (secondary)
                var sortedRows = rows.OrderBy(r => r.date).ThenBy(r => r.time).ToList();
                
                // Write sorted rows back
                for (int i = 0; i < sortedRows.Count; i++)
                {
                    int targetRow = startRow + i;
                    var sortedRow = sortedRows[i];
                    
                    for (int col = 1; col <= dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[targetRow, col];
                        cell.Value = sortedRow.values[col - 1];
                        
                        // Apply time format immediately after setting value for columns 2 and 9
                        if ((col == 2 || col == 9) && sortedRow.values[col - 1] is double)
                        {
                            cell.Style.Numberformat.Format = "h:mm";
                        }
                    }
                    
                    // Reapply date formatting to column 1
                    if (sortedRow.values[0] is DateTime)
                    {
                        worksheet.Cells[targetRow, 1].Style.Numberformat.Format = "ddd dd mmm";
                    }
                }
            }
            catch (Exception ex)
            {
                // Log error and continue without sorting (preserve original order)
                Console.WriteLine($"Errore durante l'ordinamento dei dati: {ex.Message}");
            }
        }

        /// <summary>
        /// Applies bold formatting to column headers by delegating to FormattingService.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="headerRow">The row number of the header (1-based)</param>
        public void ApplyBoldToHeaders(Sheet sheet, int headerRow)
        {
            var formattingService = new FormattingService();
            formattingService.ApplyBoldHeaders(sheet, headerRow);
        }

        /// <summary>
        /// Applies thick borders to the last row of each date group by delegating to FormattingService.
        /// Date groups are identified by comparing Data column values.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="startRow">First data row (1-based)</param>
        /// <param name="endRow">Last data row (1-based)</param>
        public void ApplyThickBordersToDateGroups(Sheet sheet, int startRow, int endRow)
        {
            var formattingService = new FormattingService();
            formattingService.ApplyDateGroupBorders(sheet, startRow, endRow, 1); // Data column is column 1
        }

        /// <summary>
        /// Saves the workbook to the specified file path.
        /// </summary>
        /// <param name="workbook">The Excel workbook to save</param>
        /// <param name="filePath">Path where the file should be saved</param>
        public void SaveWorkbook(ExcelWorkbookModel workbook, string filePath)
        {
            if (workbook == null || workbook.Package == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            try
            {
                var fileInfo = new FileInfo(filePath);
                workbook.Package.SaveAs(fileInfo);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Impossibile salvare il file: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Gets a sheet by its name.
        /// </summary>
        /// <param name="workbook">The Excel workbook</param>
        /// <param name="sheetName">The name of the sheet to retrieve</param>
        /// <returns>The sheet with the specified name, or null if not found</returns>
        public Sheet GetSheetByName(ExcelWorkbookModel workbook, string sheetName)
        {
            if (workbook == null || workbook.Package == null)
            {
                throw new ArgumentNullException(nameof(workbook));
            }

            var worksheet = workbook.Package.Workbook.Worksheets
                .FirstOrDefault(ws => ws.Name.Equals(sheetName, StringComparison.OrdinalIgnoreCase));

            return worksheet != null ? new Sheet(worksheet) : null;
        }

        /// <summary>
        /// Reads the header text from the first row of a sheet.
        /// The header is typically spread across multiple cells in the first row.
        /// Concatenates cells A1, B1, C1, and D1 to form the complete header.
        /// </summary>
        /// <param name="sheet">The sheet to read from</param>
        /// <returns>The header text from the first row</returns>
        public string ReadHeader(Sheet sheet)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var worksheet = sheet.Worksheet;
            
            // Read the first 4 cells of row 1 and concatenate them
            // Format: "DD mmm DD mmm Settimana Nreferente settimana = ..."
            var parts = new List<string>();
            
            for (int col = 1; col <= 4; col++)
            {
                var cell = worksheet.Cells[1, col];
                string cellText = cell.Text;
                
                if (!string.IsNullOrWhiteSpace(cellText))
                {
                    parts.Add(cellText);
                }
            }
            
            // Join with space, but handle the special case where "Settimana N" and "referente..."
            // should be concatenated without space (as per the expected format)
            if (parts.Count >= 3)
            {
                // Join first two parts with space: "26 gen 01 feb"
                string result = string.Join(" ", parts.Take(2));
                
                // Add third part with space: "26 gen 01 feb Settimana 5"
                result += " " + parts[2];
                
                // Add fourth part without space: "26 gen 01 feb Settimana 5referente settimana = ..."
                if (parts.Count >= 4)
                {
                    result += parts[3];
                }
                
                return result;
            }
            
            // Fallback: just join all parts with space
            return string.Join(" ", parts);
        }
    }
}
