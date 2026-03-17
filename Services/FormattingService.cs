using AuserExcelTransformer.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for applying formatting to Excel sheets.
    /// </summary>
    public interface IFormattingService
    {
        /// <summary>
        /// Applies bold formatting to all cells in the header row.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="headerRow">The row number of the header (1-indexed)</param>
        void ApplyBoldHeaders(Sheet sheet, int headerRow);

        /// <summary>
        /// Applies thick bottom borders to the last row of each date group.
        /// Date groups are identified by comparing Data column values between consecutive rows.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="dataStartRow">The first row containing data (1-indexed)</param>
        /// <param name="dataEndRow">The last row containing data (1-indexed)</param>
        /// <param name="dataColumnIndex">The column index for the Data column (1-indexed)</param>
        void ApplyDateGroupBorders(Sheet sheet, int dataStartRow, int dataEndRow, int dataColumnIndex);
    }

    /// <summary>
    /// Service for applying formatting to Excel sheets including bold headers and date group borders.
    /// </summary>
    public class FormattingService : IFormattingService
    {
        /// <summary>
        /// Applies bold formatting to all cells in the header row.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="headerRow">The row number of the header (1-indexed)</param>
        public void ApplyBoldHeaders(Sheet sheet, int headerRow)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (headerRow < 1)
            {
                throw new ArgumentException("Header row must be 1 or greater", nameof(headerRow));
            }

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            
            if (dimension == null)
            {
                return;
            }

            try
            {
                // Apply bold formatting to all cells in the header row
                for (int col = 1; col <= dimension.End.Column; col++)
                {
                    var cell = worksheet.Cells[headerRow, col];
                    cell.Style.Font.Bold = true;
                }
            }
            catch (Exception ex)
            {
                // Non-critical formatting failure - continue silently
            }
        }

        /// <summary>
        /// Applies thick bottom borders to the last row of each date group.
        /// Date groups are identified by comparing Data column values between consecutive rows.
        /// </summary>
        /// <param name="sheet">The target sheet</param>
        /// <param name="dataStartRow">The first row containing data (1-indexed)</param>
        /// <param name="dataEndRow">The last row containing data (1-indexed)</param>
        /// <param name="dataColumnIndex">The column index for the Data column (1-indexed)</param>
        public void ApplyDateGroupBorders(Sheet sheet, int dataStartRow, int dataEndRow, int dataColumnIndex)
        {
            if (sheet == null || sheet.Worksheet == null)
            {
                throw new ArgumentNullException(nameof(sheet));
            }

            if (dataStartRow < 1 || dataEndRow < dataStartRow)
            {
                throw new ArgumentException("Invalid data row range");
            }

            if (dataColumnIndex < 1)
            {
                throw new ArgumentException("Data column index must be 1 or greater", nameof(dataColumnIndex));
            }

            var worksheet = sheet.Worksheet;
            var dimension = worksheet.Dimension;
            
            if (dimension == null || dataStartRow > dataEndRow)
            {
                return;
            }

            try
            {
                // Identify date group boundaries by comparing consecutive rows
                List<int> dateGroupBoundaries = new List<int>();
                
                for (int row = dataStartRow; row <= dataEndRow; row++)
                {
                    // Check if this is the last row or if the next row has a different date
                    if (row == dataEndRow)
                    {
                        // Last row is always a boundary
                        dateGroupBoundaries.Add(row);
                    }
                    else
                    {
                        // Compare current row's date with next row's date
                        var currentDateCell = worksheet.Cells[row, dataColumnIndex];
                        var nextDateCell = worksheet.Cells[row + 1, dataColumnIndex];
                        
                        string currentDate = currentDateCell.Text?.Trim() ?? "";
                        string nextDate = nextDateCell.Text?.Trim() ?? "";
                        
                        // If dates are different, current row is the last row of a date group
                        if (!string.Equals(currentDate, nextDate, StringComparison.OrdinalIgnoreCase))
                        {
                            dateGroupBoundaries.Add(row);
                        }
                    }
                }

                // Apply thick bottom border to each date group boundary
                foreach (var boundaryRow in dateGroupBoundaries)
                {
                    for (int col = 1; col <= dimension.End.Column; col++)
                    {
                        var cell = worksheet.Cells[boundaryRow, col];
                        cell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                    }
                }
            }
            catch (Exception ex)
            {
                // Non-critical formatting failure - continue silently
            }
        }
    }
}
