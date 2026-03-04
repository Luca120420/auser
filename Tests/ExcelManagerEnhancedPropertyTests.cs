using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for ExcelManager enhanced methods.
    /// Tests universal properties across many generated inputs.
    /// </summary>
    [TestFixture]
    public class ExcelManagerEnhancedPropertyTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        // Feature: excel-output-enhancement, Property 8: Complete Row Sorting
        /// <summary>
        /// Property 8: Complete Row Sorting
        /// For any generated Excel output, all data rows SHALL be sorted first by "Data" column 
        /// in ascending order, and within each date group, rows SHALL be sorted by "Partenza" 
        /// column in ascending order.
        /// Validates: Requirements 9.1, 9.2
        /// </summary>
        [Test]
        public void Property_RowsSortedByDateThenTime()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var dateTimeGen = Arb.From(GenerateValidDateTimeRows());
            
            Prop.ForAll(dateTimeGen, rows =>
            {
                if (rows == null || rows.Count == 0)
                    return true;

                // Create a test sheet
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    var sheet = new Sheet(worksheet);

                    // Write the rows
                    _excelManager.WriteDataRowsEnhanced(sheet, rows, 3);

                    // Sort the rows
                    _excelManager.SortDataRows(sheet, 3, 3 + rows.Count - 1);

                    // Verify sorting
                    for (int i = 3; i < 3 + rows.Count - 1; i++)
                    {
                        var currentDateCell = worksheet.Cells[i, 1];
                        var nextDateCell = worksheet.Cells[i + 1, 1];
                        
                        var currentTimeCell = worksheet.Cells[i, 2];
                        var nextTimeCell = worksheet.Cells[i + 1, 2];

                        // Get dates
                        DateTime currentDate, nextDate;
                        if (!TryGetDateFromCell(currentDateCell, out currentDate) ||
                            !TryGetDateFromCell(nextDateCell, out nextDate))
                        {
                            continue; // Skip invalid dates
                        }

                        // Primary sort: dates should be in ascending order
                        if (currentDate > nextDate)
                            return false;

                        // Secondary sort: within same date, times should be in ascending order
                        if (currentDate.Date == nextDate.Date)
                        {
                            TimeSpan currentTime, nextTime;
                            if (TryGetTimeFromCell(currentTimeCell, out currentTime) &&
                                TryGetTimeFromCell(nextTimeCell, out nextTime))
                            {
                                if (currentTime > nextTime)
                                    return false;
                            }
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Generates a list of EnhancedTransformedRow with valid dates and times for testing.
        /// </summary>
        private static Gen<List<EnhancedTransformedRow>> GenerateValidDateTimeRows()
        {
            return Gen.Choose(1, 20).SelectMany(count =>
            {
                var rowGens = new List<Gen<EnhancedTransformedRow>>();
                
                for (int i = 0; i < count; i++)
                {
                    rowGens.Add(GenerateSingleRow());
                }

                return Gen.Sequence(rowGens).Select(rows => rows.ToList());
            });
        }

        /// <summary>
        /// Generates a single EnhancedTransformedRow with valid date and time.
        /// </summary>
        private static Gen<EnhancedTransformedRow> GenerateSingleRow()
        {
            var dateGen = Gen.Choose(0, 365).Select(days =>
            {
                var date = DateTime.Today.AddDays(days - 180); // ±6 months from today
                return date.ToString("dd/MM/yyyy");
            });

            var timeGen = Gen.Choose(0, 1439).Select(minutes =>
            {
                var hours = minutes / 60;
                var mins = minutes % 60;
                return $"{hours:D2}:{mins:D2}";
            });

            var nameGen = Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Anna", "Neri Paolo", "Gialli Maria");

            return Gen.Zip(dateGen, timeGen, nameGen).Select(tuple =>
            {
                var (date, time, name) = tuple;
                return new EnhancedTransformedRow
                {
                    Data = date,
                    Partenza = time,
                    Assistito = name,
                    Indirizzo = "Via Roma 1",
                    Destinazione = "Ospedale",
                    Note = "",
                    Auto = "Auto1",
                    Volontario = "Vol1",
                    Arrivo = "",
                    Avv = "",
                    IndirizzoGasnet = "Via Roma 1, Milano",
                    NoteGasnet = ""
                };
            });
        }

        /// <summary>
        /// Tries to extract a DateTime from an Excel cell.
        /// </summary>
        private bool TryGetDateFromCell(ExcelRange cell, out DateTime date)
        {
            date = DateTime.MinValue;

            if (cell.Value is DateTime)
            {
                date = (DateTime)cell.Value;
                return true;
            }

            if (cell.Value is double)
            {
                try
                {
                    date = DateTime.FromOADate((double)cell.Value);
                    return true;
                }
                catch
                {
                    return false;
                }
            }

            if (cell.Value is string)
            {
                return DateTime.TryParse((string)cell.Value, out date);
            }

            return false;
        }

        /// <summary>
        /// Tries to extract a TimeSpan from an Excel cell.
        /// </summary>
        private bool TryGetTimeFromCell(ExcelRange cell, out TimeSpan time)
        {
            time = TimeSpan.Zero;

            if (cell.Value == null)
                return false;

            var valueStr = cell.Value.ToString();
            
            if (TimeSpan.TryParse(valueStr, out time))
                return true;

            // Try parsing HH:mm format
            if (valueStr.Contains(":"))
            {
                var parts = valueStr.Split(':');
                if (parts.Length >= 2 &&
                    int.TryParse(parts[0], out int hours) &&
                    int.TryParse(parts[1], out int minutes))
                {
                    time = new TimeSpan(hours, minutes, 0);
                    return true;
                }
            }

            return false;
        }

        // Feature: excel-output-enhancement, Property 10: Round-Trip Data Integrity
        /// <summary>
        /// Property 10: Round-Trip Data Integrity
        /// For any valid input data, writing it to Excel and then reading it back SHALL produce 
        /// a dataset equivalent to the original transformed data (excluding formatting).
        /// Validates: Requirements 11.1
        /// </summary>
        [Test]
        public void Property_RoundTripPreservesData()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var rowsGen = Arb.From(GenerateValidDateTimeRows());
            
            Prop.ForAll(rowsGen, originalRows =>
            {
                if (originalRows == null || originalRows.Count == 0)
                    return true;

                // Write to Excel
                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    var sheet = new Sheet(worksheet);

                    _excelManager.WriteDataRowsEnhanced(sheet, originalRows, 3);

                    // Read back from Excel
                    var readBackRows = new List<EnhancedTransformedRow>();
                    
                    for (int row = 3; row < 3 + originalRows.Count; row++)
                    {
                        var readRow = new EnhancedTransformedRow
                        {
                            Data = GetCellValueAsString(worksheet.Cells[row, 1]),
                            Partenza = GetCellValueAsString(worksheet.Cells[row, 2]),
                            Assistito = GetCellValueAsString(worksheet.Cells[row, 3]),
                            Indirizzo = GetCellValueAsString(worksheet.Cells[row, 4]),
                            Destinazione = GetCellValueAsString(worksheet.Cells[row, 5]),
                            Note = GetCellValueAsString(worksheet.Cells[row, 6]),
                            Auto = GetCellValueAsString(worksheet.Cells[row, 7]),
                            Volontario = GetCellValueAsString(worksheet.Cells[row, 8]),
                            Arrivo = GetCellValueAsString(worksheet.Cells[row, 9]),
                            Avv = GetCellValueAsString(worksheet.Cells[row, 10]),
                            IndirizzoGasnet = GetCellValueAsString(worksheet.Cells[row, 11]),
                            NoteGasnet = GetCellValueAsString(worksheet.Cells[row, 12])
                        };
                        
                        readBackRows.Add(readRow);
                    }

                    // Compare data (excluding date formatting differences)
                    if (originalRows.Count != readBackRows.Count)
                        return false;

                    for (int i = 0; i < originalRows.Count; i++)
                    {
                        if (!DataEquivalent(originalRows[i], readBackRows[i]))
                            return false;
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Gets the cell value as a string, handling dates and other types.
        /// </summary>
        private string GetCellValueAsString(ExcelRange cell)
        {
            if (cell.Value == null)
                return string.Empty;

            // Handle DateTime values (dates are stored as DateTime in Excel)
            if (cell.Value is DateTime dt)
            {
                // Return in dd/MM/yyyy format to match input
                return dt.ToString("dd/MM/yyyy");
            }

            return cell.Value.ToString() ?? string.Empty;
        }

        /// <summary>
        /// Compares two EnhancedTransformedRow objects for data equivalence.
        /// Ignores formatting differences (e.g., date format variations).
        /// </summary>
        private bool DataEquivalent(EnhancedTransformedRow original, EnhancedTransformedRow readBack)
        {
            // Compare dates (normalize format)
            if (!DatesEquivalent(original.Data, readBack.Data))
                return false;

            // Compare all other fields (exact match)
            return original.Partenza == readBack.Partenza &&
                   original.Assistito == readBack.Assistito &&
                   original.Indirizzo == readBack.Indirizzo &&
                   original.Destinazione == readBack.Destinazione &&
                   original.Note == readBack.Note &&
                   original.Auto == readBack.Auto &&
                   original.Volontario == readBack.Volontario &&
                   original.Arrivo == readBack.Arrivo &&
                   original.Avv == readBack.Avv &&
                   original.IndirizzoGasnet == readBack.IndirizzoGasnet &&
                   original.NoteGasnet == readBack.NoteGasnet;
        }

        /// <summary>
        /// Compares two date strings for equivalence, handling format variations.
        /// </summary>
        private bool DatesEquivalent(string date1, string date2)
        {
            if (string.IsNullOrWhiteSpace(date1) && string.IsNullOrWhiteSpace(date2))
                return true;

            DateTime dt1, dt2;
            
            if (DateTime.TryParseExact(date1, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt1) &&
                DateTime.TryParseExact(date2, new[] { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd" },
                CultureInfo.InvariantCulture, DateTimeStyles.None, out dt2))
            {
                return dt1.Date == dt2.Date;
            }

            if (DateTime.TryParse(date1, out dt1) && DateTime.TryParse(date2, out dt2))
            {
                return dt1.Date == dt2.Date;
            }

            return date1 == date2;
        }
    }
}
