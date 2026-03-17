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

                    // Find the position to insert the new sheet
                    // It should be right after the last numbered sheet (the one it's based on: n+1 after n)
                    var worksheets = workbook.Package.Workbook.Worksheets;
                    int insertPosition = 0; // Default: add at the beginning
                    
                    // Find the position of the sheet with number (sheetNumber - 1)
                    // The new sheet should be inserted right after it
                    string previousSheetName = (sheetNumber - 1).ToString();
                    
                    for (int i = 0; i < worksheets.Count; i++)
                    {
                        string sheetName = worksheets[i].Name;
                        if (sheetName.Equals(previousSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            // Found the previous numbered sheet, insert after it
                            insertPosition = i + 1;
                            break;
                        }
                    }
                    
                    // If we didn't find the previous sheet, find the last numbered sheet
                    if (insertPosition == 0)
                    {
                        int lastNumberedSheetPosition = -1;
                        
                        for (int i = 0; i < worksheets.Count; i++)
                        {
                            string sheetName = worksheets[i].Name;
                            // Check if this is a numbered sheet (1-53 range)
                            if (int.TryParse(sheetName.Trim(), out int number))
                            {
                                if (number >= 1 && number <= 53)
                                {
                                    lastNumberedSheetPosition = i;
                                }
                            }
                        }
                        
                        // Insert after the last numbered sheet, or at the beginning if none found
                        insertPosition = lastNumberedSheetPosition + 1;
                    }

                    // Insert the new sheet at the calculated position
                    var worksheet = worksheets.Add(sheetNumber.ToString());

                    // Move the sheet to the correct position if needed
                    if (insertPosition < worksheets.Count - 1)
                    {
                        // The sheet was added at the end, now move it to the correct position
                        // EPPlus MoveBefore/MoveAfter methods use 0-based indexing
                        worksheets.MoveBefore(worksheets.Count - 1, insertPosition);
                    }

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
        /// <param name="targetWeekMonday">The Monday date of the target week (from cell A1)</param>
        public void AppendFissiData(Sheet targetSheet, Sheet fissiSheet, int startRow, DateTime targetWeekMonday)
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
                // Skip rows where column 1 (Data) is empty or null
                var dataCell = fissiWorksheet.Cells[sourceRow, 1];
                if (dataCell.Value == null || string.IsNullOrWhiteSpace(dataCell.Text))
                {
                    continue; // Skip this row
                }

                // Copy mapped columns from fissi sheet
                for (int fissiCol = 1; fissiCol <= Math.Min(9, fissiDimension.End.Column); fissiCol++)
                {
                    int targetCol = columnMapping[fissiCol - 1];
                    
                    var sourceCell = fissiWorksheet.Cells[sourceRow, fissiCol];
                    var targetCell = targetWorksheet.Cells[targetRow, targetCol];

                    // Special handling for Data column (1) - calculate next week's same day
                    if (fissiCol == 1)
                    {
                        var sourceValue = sourceCell.Value;
                        DateTime sourceDate;
                        
                        // Parse the source date
                        if (sourceValue is DateTime dt)
                        {
                            sourceDate = dt;
                        }
                        else if (sourceValue is double)
                        {
                            sourceDate = DateTime.FromOADate((double)sourceValue);
                        }
                        else if (DateTime.TryParse(sourceValue?.ToString(), out DateTime parsedDate))
                        {
                            sourceDate = parsedDate;
                        }
                        else
                        {
                            // Can't parse - copy as-is
                            targetCell.Value = sourceValue;
                            continue;
                        }
                        
                        // Calculate the same day of week within the target week range
                        DateTime nextWeekDate = CalculateNextWeekSameDay(sourceDate, targetWeekMonday);
                        targetCell.Value = nextWeekDate;
                        targetCell.Style.Numberformat.Format = "ddd dd mmm";
                    }
                    // Col 4 (Indirizzo) - write VLOOKUP formula instead of static value
                    else if (fissiCol == 4)
                    {
                        targetCell.Formula = $"VLOOKUP(C{targetRow},assistiti!A:C,2,FALSE)";
                    }
                    // Col 6 (Note) - write VLOOKUP formula instead of static value
                    else if (fissiCol == 6)
                    {
                        targetCell.Formula = $"VLOOKUP(C{targetRow},assistiti!A:C,3,FALSE)";
                    }
                    // Special handling for time columns (2 = Partenza, 9 = Arrivo)
                    else if (fissiCol == 2 || fissiCol == 9)
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
                    // Skip formatting for formula cells (col 4 and 6)
                    if (sourceCell.Style != null && fissiCol != 4 && fissiCol != 6)
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
                // Apply thin borders to all columns in the row (including empty cells)
                for (int c = 1; c <= 12; c++)
                {
                    var borderCell = targetWorksheet.Cells[targetRow, c];
                    borderCell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                // Apply wrap text to Assistito (col 3) for all fissi rows
                targetWorksheet.Cells[targetRow, 3].Style.WrapText = true;
                targetRow++;
            }
        }

        /// <summary>
        /// Appends data from the laboratori sheet to the target sheet.
        /// Laboratori sheet has 10 columns (same as fissi plus an additional "Avv" column).
        /// </summary>
        public void AppendLaboratoriData(Sheet targetSheet, Sheet laboratoriSheet, int startRow, DateTime targetWeekMonday)
        {
            if (targetSheet == null || targetSheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(targetSheet));
            }

            if (laboratoriSheet == null || laboratoriSheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(laboratoriSheet));
            }

            var laboratoriWorksheet = laboratoriSheet.Worksheet;
            var targetWorksheet = targetSheet.Worksheet;

            // Get the dimensions of the laboratori sheet
            var laboratoriDimension = laboratoriWorksheet.Dimension;
            if (laboratoriDimension == null)
            {
                return; // Empty sheet
            }

            // Detect where data starts in laboratori sheet
            // Check if row 2 contains "Data" in cell A2 (column header)
            int laboratoriDataStartRow = 3; // Default: skip rows 1 and 2

            var row2Cell = laboratoriWorksheet.Cells[2, 1];
            string row2Text = row2Cell.Text?.Trim() ?? string.Empty;

            // If row 2 contains "Data", it's a column header row, so skip it and start from row 3
            if (row2Text.Equals("Data", StringComparison.OrdinalIgnoreCase))
            {
                laboratoriDataStartRow = 3;
            }
            else
            {
                // If row 2 doesn't contain "Data", check row 1
                var row1Cell = laboratoriWorksheet.Cells[1, 1];
                string row1Text = row1Cell.Text?.Trim() ?? string.Empty;

                if (row1Text.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    // Row 1 is column headers, start from row 2
                    laboratoriDataStartRow = 2;
                }
                else
                {
                    // Neither row 1 nor row 2 contains "Data", assume row 2 is data
                    laboratoriDataStartRow = 2;
                }
            }

            // Column mapping from laboratori (10 cols) to new structure (14 cols):
            // Laboratori Col 1 (Data) → Target Col 1 (Data)
            // Laboratori Col 2 (Partenza) → Target Col 2 (Ora Inizio Servizio)
            // Laboratori Col 3 (Assistito) → Target Col 3 (Assistito)
            // Laboratori Col 4 (Indirizzo) → Target Col 4 (Indirizzo)
            // Laboratori Col 5 (Destinazione) → Target Col 5 (Destinazione)
            // Laboratori Col 6 (Note) → Target Col 6 (Note)
            // Laboratori Col 7 (Auto) → Target Col 7 (Auto)
            // Laboratori Col 8 (Volontario) → Target Col 8 (Volontario)
            // Laboratori Col 9 (Arrivo) → Target Col 9 (Arrivo)
            // Laboratori Col 10 (Avv) → Target Col 10 (Avv)
            // Target Cols 11-14 remain empty for other data sources

            int[] columnMapping = new int[] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10 };

            // Copy data rows only (skip header rows)
            int targetRow = startRow;
            for (int sourceRow = laboratoriDataStartRow; sourceRow <= laboratoriDimension.End.Row; sourceRow++)
            {
                // Skip rows where column 1 (Data) is empty or null
                var dataCell = laboratoriWorksheet.Cells[sourceRow, 1];
                if (dataCell.Value == null || string.IsNullOrWhiteSpace(dataCell.Text))
                {
                    continue; // Skip this row
                }

                // Copy mapped columns from laboratori sheet
                for (int laboratoriCol = 1; laboratoriCol <= Math.Min(10, laboratoriDimension.End.Column); laboratoriCol++)
                {
                    int targetCol = columnMapping[laboratoriCol - 1];

                    var sourceCell = laboratoriWorksheet.Cells[sourceRow, laboratoriCol];
                    var targetCell = targetWorksheet.Cells[targetRow, targetCol];

                    // Special handling for Data column (1) - calculate next week's same day
                    if (laboratoriCol == 1)
                    {
                        var sourceValue = sourceCell.Value;
                        DateTime sourceDate;
                        
                        // Parse the source date
                        if (sourceValue is DateTime dt)
                        {
                            sourceDate = dt;
                        }
                        else if (sourceValue is double)
                        {
                            sourceDate = DateTime.FromOADate((double)sourceValue);
                        }
                        else if (DateTime.TryParse(sourceValue?.ToString(), out DateTime parsedDate))
                        {
                            sourceDate = parsedDate;
                        }
                        else
                        {
                            // Can't parse - copy as-is
                            targetCell.Value = sourceValue;
                            continue;
                        }
                        
                        // Calculate the same day of week within the target week range
                        DateTime nextWeekDate = CalculateNextWeekSameDay(sourceDate, targetWeekMonday);
                        targetCell.Value = nextWeekDate;
                        targetCell.Style.Numberformat.Format = "ddd dd mmm";
                    }
                    // Col 4 (Indirizzo) - copy static value from laboratori sheet (no VLOOKUP)
                    else if (laboratoriCol == 4)
                    {
                        targetCell.Value = sourceCell.Value;
                    }
                    // Col 6 (Note) - copy static value from laboratori sheet (no VLOOKUP)
                    else if (laboratoriCol == 6)
                    {
                        targetCell.Value = sourceCell.Value;
                    }
                    // Special handling for time columns (2 = Partenza, 9 = Arrivo)
                    else if (laboratoriCol == 2 || laboratoriCol == 9)
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
                        if (!string.IsNullOrEmpty(sourceCell.Style.Numberformat.Format) && laboratoriCol != 2 && laboratoriCol != 9)
                        {
                            targetCell.Style.Numberformat.Format = sourceCell.Style.Numberformat.Format;
                        }

                        // Always apply time format to columns 2 (Partenza) and 9 (Arrivo)
                        if (laboratoriCol == 2 || laboratoriCol == 9)
                        {
                            targetCell.Style.Numberformat.Format = "h:mm";
                        }
                    }

                    // Laboratori rows: Assistito (col 3) and Indirizzo (col 4) must be Italic + Tahoma + size 9
                    // Applied after formatting copy so it is never overwritten
                    if (targetCol == 3 || targetCol == 4)
                    {
                        targetCell.Style.Font.Italic = true;
                        targetCell.Style.Font.Name = "Tahoma";
                        targetCell.Style.Font.Size = 9;
                    }
                }
                // Apply thin borders to all columns in the row (including empty cells)
                for (int c = 1; c <= 12; c++)
                {
                    var borderCell = targetWorksheet.Cells[targetRow, c];
                    borderCell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }
                // Apply wrap text to Assistito (col 3) for all laboratori rows
                targetWorksheet.Cells[targetRow, 3].Style.WrapText = true;
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
                var assistitoCell = sheet.Worksheet.Cells[currentRow, col++];
                assistitoCell.Value = row.Assistito;
                assistitoCell.Style.WrapText = true;
                sheet.Worksheet.Cells[currentRow, col++].Formula = $"VLOOKUP(C{currentRow},assistiti!A:C,2,FALSE)";
                
                // Column 5: Destinazione
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Destinazione;
                
                // Column 6: Note - VLOOKUP formula
                sheet.Worksheet.Cells[currentRow, col++].Formula = $"VLOOKUP(C{currentRow},assistiti!A:C,3,FALSE)";
                
                // Column 7: Auto
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Auto;
                
                // Column 8: Volontario
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Volontario;
                
                // Column 9: Arrivo
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Arrivo;
                
                // Column 10: Avv (from fissi lookup)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.Avv;
                
                // Column 11: Indirizzo Gasnet
                sheet.Worksheet.Cells[currentRow, col++].Value = row.IndirizzoGasnet;
                
                // Column 12: Note Gasnet (from CSV)
                sheet.Worksheet.Cells[currentRow, col++].Value = row.NoteGasnet;

                // Apply thin borders to all columns in the row (including empty cells)
                for (int c = 1; c <= 12; c++)
                {
                    var borderCell = sheet.Worksheet.Cells[currentRow, c];
                    borderCell.Style.Border.Top.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Bottom.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Left.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                    borderCell.Style.Border.Right.Style = OfficeOpenXml.Style.ExcelBorderStyle.Thin;
                }

                // Apply yellow highlight if flagged
                if (row.IsYellow)
                {
                    var dimension = sheet.Worksheet.Dimension;
                    int maxCol = dimension?.End.Column ?? 12;
                    for (int c = 1; c <= maxCol; c++)
                    {
                        var cell = sheet.Worksheet.Cells[currentRow, c];
                        cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                    }
                }

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
                // Each row stores values, formulas, and per-cell font styles so they survive the sort
                var rows = new List<(int rowNum, DateTime date, TimeSpan time, object[] values, string[] formulas, bool[] italic, string[] fontName, float[] fontSize, bool[] wrapText, string[] fillColor, OfficeOpenXml.Style.ExcelBorderStyle[] borderStyle)>();
                
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
                    
                    // Read all cell values, formulas, and font styles for this row
                    var values = new object[dimension.End.Column];
                    var formulas = new string[dimension.End.Column];
                    var italic = new bool[dimension.End.Column];
                    var fontName = new string[dimension.End.Column];
                    var fontSize = new float[dimension.End.Column];
                    var wrapText = new bool[dimension.End.Column];
                    var fillColor = new string[dimension.End.Column];
                    var borderStyle = new OfficeOpenXml.Style.ExcelBorderStyle[dimension.End.Column];
                    for (int col = 1; col <= dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[row, col];
                        formulas[col - 1] = cell.Formula ?? string.Empty;
                        values[col - 1] = cell.Value;
                        italic[col - 1] = cell.Style.Font.Italic;
                        fontName[col - 1] = cell.Style.Font.Name;
                        fontSize[col - 1] = cell.Style.Font.Size;
                        wrapText[col - 1] = cell.Style.WrapText;
                        // Only capture yellow fill
                        bool cellIsYellow = cell.Style.Fill.PatternType == OfficeOpenXml.Style.ExcelFillStyle.Solid
                            && cell.Style.Fill.BackgroundColor.Rgb == "FFFFFF00";
                        fillColor[col - 1] = cellIsYellow ? "FFFFFF00" : null;
                        borderStyle[col - 1] = cell.Style.Border.Top.Style;
                    }
                    
                    rows.Add((row, date, time, values, formulas, italic, fontName, fontSize, wrapText, fillColor, borderStyle));
                }
                
                // Sort by date (primary) then time (secondary)
                var sortedRows = rows.OrderBy(r => r.date).ThenBy(r => r.time).ToList();
                
                // Write sorted rows back — restore formulas where present, values otherwise
                for (int i = 0; i < sortedRows.Count; i++)
                {
                    int targetRow = startRow + i;
                    var sortedRow = sortedRows[i];
                    
                    for (int col = 1; col <= dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[targetRow, col];
                        var formula = sortedRow.formulas[col - 1];

                        if (!string.IsNullOrEmpty(formula))
                        {
                            // Rewrite formula — update row references to the new target row
                            // Replace row numbers in cell references (e.g. C3 → C{targetRow})
                            var updatedFormula = System.Text.RegularExpressions.Regex.Replace(
                                formula,
                                @"(?<=[A-Z])\d+",
                                targetRow.ToString());
                            cell.Formula = updatedFormula;
                        }
                        else
                        {
                            cell.Value = sortedRow.values[col - 1];
                            
                            // Apply time format immediately after setting value for columns 2 and 9
                            if ((col == 2 || col == 9) && sortedRow.values[col - 1] is double)
                            {
                                cell.Style.Numberformat.Format = "h:mm";
                            }
                        }

                        // Restore font italic, name, size, and wrap text
                        cell.Style.Font.Italic = sortedRow.italic[col - 1];
                        if (!string.IsNullOrEmpty(sortedRow.fontName[col - 1]))
                        {
                            cell.Style.Font.Name = sortedRow.fontName[col - 1];
                        }
                        if (sortedRow.fontSize[col - 1] > 0)
                        {
                            cell.Style.Font.Size = sortedRow.fontSize[col - 1];
                        }
                        cell.Style.WrapText = sortedRow.wrapText[col - 1];

                        // Restore fill color — always explicitly set to avoid EPPlus shared style bleed
                        if (sortedRow.fillColor[col - 1] == "FFFFFF00")
                        {
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Yellow);
                        }
                        else
                        {
                            // Explicitly set no fill — use Solid white to force a distinct style object
                            cell.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            cell.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.White);
                        }

                        // Restore borders
                        var bs = sortedRow.borderStyle[col - 1];
                        cell.Style.Border.Top.Style = bs;
                        cell.Style.Border.Bottom.Style = bs;
                        cell.Style.Border.Left.Style = bs;
                        cell.Style.Border.Right.Style = bs;
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

                    // Trim whitespace from the search name for more robust matching
                    string trimmedSheetName = sheetName?.Trim() ?? string.Empty;

                    var worksheet = workbook.Package.Workbook.Worksheets
                        .FirstOrDefault(ws => ws.Name.Trim().Equals(trimmedSheetName, StringComparison.OrdinalIgnoreCase));

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

        /// <summary>
        /// Identifies the column index containing "Volontario" in the header row.
        /// Performs case-insensitive search for the column header.
        /// </summary>
        /// <param name="sheet">The sheet to search</param>
        /// <returns>The 1-based column index of the Volontario column</returns>
        /// <exception cref="InvalidOperationException">If the Volontario column is not found</exception>
        public int GetVolontarioColumnIndex(Sheet sheet)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            var worksheet = sheet.Worksheet;
            
            // Determine the header row - typically row 1, but we'll search the first few rows
            // to be more flexible
            int headerRow = 1;
            
            // Get the dimensions of the worksheet to know how many columns to check
            var dimension = worksheet.Dimension;
            if (dimension == null)
            {
                throw new InvalidOperationException("Il foglio è vuoto.");
            }

            int maxColumn = dimension.End.Column;
            
            // Search for "Volontario" in the header row (case-insensitive)
            for (int col = 1; col <= maxColumn; col++)
            {
                var cell = worksheet.Cells[headerRow, col];
                string cellText = cell.Text;
                
                if (!string.IsNullOrWhiteSpace(cellText) && 
                    cellText.IndexOf("Volontario", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return col;
                }
            }
            
            // If not found in row 1, try row 2 (some sheets might have a different structure)
            headerRow = 2;
            for (int col = 1; col <= maxColumn; col++)
            {
                var cell = worksheet.Cells[headerRow, col];
                string cellText = cell.Text;
                
                if (!string.IsNullOrWhiteSpace(cellText) && 
                    cellText.IndexOf("Volontario", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return col;
                }
            }
            
            throw new InvalidOperationException("La colonna 'Volontario' non è stata trovata nel foglio selezionato.");
        }
        /// <summary>
                /// Reads all data rows from a sheet and returns them as a list of dictionaries.
                /// Each dictionary maps column names to cell values for that row.
                /// </summary>
                /// <param name="sheet">The sheet to read from</param>
                /// <returns>List of dictionaries where each dictionary represents a row with column name -> cell value mappings</returns>
                /// <exception cref="ArgumentNullException">If sheet is null</exception>
                /// <exception cref="InvalidOperationException">If sheet is empty or has no header row</exception>
                public List<Dictionary<string, string>> ReadAllRowsWithColumnNames(Sheet sheet)
                {
                    if (sheet == null || sheet.Worksheet == null)
                    {
                        throw new ArgumentNullException(nameof(sheet));
                    }

                    var worksheet = sheet.Worksheet;
                    var dimension = worksheet.Dimension;

                    if (dimension == null)
                    {
                        throw new InvalidOperationException("Il foglio è vuoto.");
                    }

                    var result = new List<Dictionary<string, string>>();

                    // Determine header row by looking for "Volontario" column
                    int headerRow = 1;
                    int maxColumn = dimension.End.Column;
                    int maxRow = dimension.End.Row;
                    
                    // Check if row 1 has "Volontario" column
                    bool foundInRow1 = false;
                    for (int col = 1; col <= maxColumn; col++)
                    {
                        var cell = worksheet.Cells[1, col];
                        string cellText = cell.Text?.Trim() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(cellText) && 
                            cellText.IndexOf("Volontario", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            foundInRow1 = true;
                            break;
                        }
                    }
                    
                    // If not found in row 1, check row 2
                    if (!foundInRow1)
                    {
                        bool foundInRow2 = false;
                        for (int col = 1; col <= maxColumn; col++)
                        {
                            var cell = worksheet.Cells[2, col];
                            string cellText = cell.Text?.Trim() ?? string.Empty;
                            if (!string.IsNullOrWhiteSpace(cellText) && 
                                cellText.IndexOf("Volontario", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                foundInRow2 = true;
                                break;
                            }
                        }
                        
                        if (foundInRow2)
                        {
                            headerRow = 2;
                        }
                    }

                    // Read column names from header row
                    var columnNames = new Dictionary<int, string>();
                    for (int col = 1; col <= maxColumn; col++)
                    {
                        var cell = worksheet.Cells[headerRow, col];
                        string columnName = cell.Text?.Trim() ?? $"Column{col}";
                        columnNames[col] = columnName;
                    }

                    // Read all data rows (starting from row after header)
                    for (int row = headerRow + 1; row <= maxRow; row++)
                    {
                        var rowData = new Dictionary<string, string>();
                        bool hasData = false;

                        for (int col = 1; col <= maxColumn; col++)
                        {
                            var cell = worksheet.Cells[row, col];
                            string cellValue = cell.Text?.Trim() ?? string.Empty;

                            if (!string.IsNullOrWhiteSpace(cellValue))
                            {
                                hasData = true;
                            }

                            rowData[columnNames[col]] = cellValue;
                        }

                        // Only add rows that have at least some data
                        if (hasData)
                        {
                            result.Add(rowData);
                        }
                    }

                    return result;
                }

                /// <summary>
                /// Identifies volunteer assignments by matching volunteer surnames to rows in the sheet.
                /// For each volunteer, finds all rows where the volunteer's surname appears in the Volontario column
                /// using case-insensitive substring matching.
                /// </summary>
                /// <param name="sheet">The sheet to read from</param>
                /// <param name="volunteers">Dictionary mapping volunteer surnames to email addresses</param>
                /// <returns>List of VolunteerAssignment objects containing surname, email, and assigned rows</returns>
                /// <exception cref="ArgumentNullException">If sheet or volunteers is null</exception>
                /// <exception cref="InvalidOperationException">If Volontario column is not found</exception>
                public List<VolunteerAssignment> IdentifyVolunteerAssignments(Sheet sheet, Dictionary<string, string> volunteers)
                {
                    if (sheet == null || sheet.Worksheet == null)
                    {
                        throw new ArgumentNullException(nameof(sheet));
                    }

                    if (volunteers == null)
                    {
                        throw new ArgumentNullException(nameof(volunteers));
                    }

                    var result = new List<VolunteerAssignment>();

                    // Get the Volontario column index
                    int volontarioColumnIndex = GetVolontarioColumnIndex(sheet);

                    // Read all rows with column names
                    var allRows = ReadAllRowsWithColumnNames(sheet);

                    // Determine the header row (same logic as ReadAllRowsWithColumnNames)
                    var worksheet = sheet.Worksheet;
                    var dimension = worksheet.Dimension;
                    int headerRow = 1;
                    int maxColumn = dimension.End.Column;
                    
                    // Check if row 1 has "Volontario" column
                    bool foundInRow1 = false;
                    for (int col = 1; col <= maxColumn; col++)
                    {
                        var cell = worksheet.Cells[1, col];
                        string cellText = cell.Text?.Trim() ?? string.Empty;
                        if (!string.IsNullOrWhiteSpace(cellText) && 
                            cellText.IndexOf("Volontario", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            foundInRow1 = true;
                            break;
                        }
                    }
                    
                    // If not found in row 1, check row 2
                    if (!foundInRow1)
                    {
                        headerRow = 2;
                    }

                    // Get the column name for the Volontario column from the correct header row
                    var volontarioColumnName = worksheet.Cells[headerRow, volontarioColumnIndex].Text?.Trim() ?? $"Column{volontarioColumnIndex}";

                    // For each volunteer, find matching rows
                    foreach (var volunteer in volunteers)
                    {
                        string surname = volunteer.Key;
                        string email = volunteer.Value;

                        var assignedRows = new List<Dictionary<string, string>>();

                        // Find all rows where the surname appears in the Volontario column
                        foreach (var row in allRows)
                        {
                            // Check if this row has the Volontario column
                            if (row.ContainsKey(volontarioColumnName))
                            {
                                string volontarioValue = row[volontarioColumnName];

                                // Case-insensitive substring matching
                                if (!string.IsNullOrWhiteSpace(volontarioValue) &&
                                    volontarioValue.IndexOf(surname, StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    assignedRows.Add(row);
                                }
                            }
                        }

                        // Only add volunteers who have at least one assigned row
                        if (assignedRows.Count > 0)
                        {
                            result.Add(new VolunteerAssignment
                            {
                                Surname = surname,
                                Email = email,
                                AssignedRows = assignedRows
                            });
                        }
                    }

                    return result;
                }

        /// <summary>
        /// Calculates the next occurrence of the same day of the week from today's date.
        /// For example, if the source date is a Friday (23/01/2026) and the target week starts on Monday 02/02/2026,
        /// this returns Friday 06/02/2026 (the Friday within that week range).
        /// </summary>
        /// <param name="sourceDate">The source date from fissi/laboratori sheet</param>
        /// <param name="targetWeekMonday">The Monday date (from cell A1) that starts the target week</param>
        /// <returns>The date with the same day of week within the target week</returns>
        private DateTime CalculateNextWeekSameDay(DateTime sourceDate, DateTime targetWeekMonday)
        {
            DayOfWeek sourceDayOfWeek = sourceDate.DayOfWeek;
            
            // Calculate how many days from Monday to the target day of week
            // Monday = 1, Tuesday = 2, ..., Sunday = 0
            int daysFromMonday = ((int)sourceDayOfWeek - (int)DayOfWeek.Monday + 7) % 7;
            
            // Add those days to the target week's Monday to get the correct date
            return targetWeekMonday.AddDays(daysFromMonday);
        }

    }
}
