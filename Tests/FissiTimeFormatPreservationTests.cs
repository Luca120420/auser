using System;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Preservation property tests for fissi time format fix.
    /// 
    /// IMPORTANT: These tests verify behaviors that should remain UNCHANGED.
    /// These tests should PASS on unfixed code and continue to PASS after the fix.
    /// 
    /// This follows the observation-first methodology:
    /// 1. Observe behavior on UNFIXED code for non-buggy inputs
    /// 2. Write property-based tests capturing observed behavior patterns
    /// 3. Run tests on UNFIXED code - they should PASS
    /// 4. After fix is implemented, re-run tests - they should still PASS
    /// 
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// </summary>
    [TestFixture]
    public class FissiTimeFormatPreservationTests
    {
        private ExcelManager _excelManager = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Property 2: Preservation - Non-Time Columns (1, 3, 4, 5, 6, 7, 8) Preserve Values
        /// 
        /// For any fissi sheet data where non-time columns (NOT 2 or 9) contain values,
        /// the AppendFissiData method SHALL continue to copy these values correctly
        /// without modification.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1**
        /// </summary>
        [Test]
        public void Property_Preservation_NonTimeColumnsPreserveValues()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                ArbitraryNonTimeColumnData(),
                (NonTimeColumnData data) =>
                {
                    // Arrange - Create fissi sheet with data in non-time columns
                    using (var package = new ExcelPackage())
                    {
                        var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                        var targetWorksheet = package.Workbook.Worksheets.Add("target");

                        // Add header row
                        fissiWorksheet.Cells[2, 1].Value = "Data";
                        fissiWorksheet.Cells[2, 3].Value = "Assistito";
                        fissiWorksheet.Cells[2, 4].Value = "Indirizzo";
                        fissiWorksheet.Cells[2, 5].Value = "Destinazione";
                        fissiWorksheet.Cells[2, 6].Value = "Note";
                        fissiWorksheet.Cells[2, 7].Value = "Auto";
                        fissiWorksheet.Cells[2, 8].Value = "Volontario";

                        // Add data row with values in non-time columns (1, 3, 4, 5, 6, 7, 8)
                        fissiWorksheet.Cells[3, 1].Value = data.Data;
                        fissiWorksheet.Cells[3, 3].Value = data.Assistito;
                        fissiWorksheet.Cells[3, 4].Value = data.Indirizzo;
                        fissiWorksheet.Cells[3, 5].Value = data.Destinazione;
                        fissiWorksheet.Cells[3, 6].Value = data.Note;
                        fissiWorksheet.Cells[3, 7].Value = data.Auto;
                        fissiWorksheet.Cells[3, 8].Value = data.Volontario;

                        var fissiSheet = new Sheet(fissiWorksheet);
                        var targetSheet = new Sheet(targetWorksheet);

                        // Act
                        _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                        // Assert - Verify non-time columns preserve their values
                        var dataValue = targetWorksheet.Cells[1, 1].Value?.ToString();
                        var assistitoValue = targetWorksheet.Cells[1, 3].Value?.ToString();
                        var indirizzoValue = targetWorksheet.Cells[1, 4].Value?.ToString();
                        var destinazioneValue = targetWorksheet.Cells[1, 5].Value?.ToString();
                        var noteValue = targetWorksheet.Cells[1, 6].Value?.ToString();
                        var autoValue = targetWorksheet.Cells[1, 7].Value?.ToString();
                        var volontarioValue = targetWorksheet.Cells[1, 8].Value?.ToString();

                        return dataValue == data.Data &&
                               assistitoValue == data.Assistito &&
                               indirizzoValue == data.Indirizzo &&
                               destinazioneValue == data.Destinazione &&
                               noteValue == data.Note &&
                               autoValue == data.Auto &&
                               volontarioValue == data.Volontario;
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 2: Preservation - Text Values in Any Column Preserve Text Format
        /// 
        /// For any fissi sheet data where columns contain text values (including columns 2 and 9),
        /// the AppendFissiData method SHALL continue to preserve the text format and content.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.2**
        /// </summary>
        [Test]
        public void Property_Preservation_TextValuesPreserveFormat()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            Prop.ForAll(
                ArbitraryTextColumnData(),
                (TextColumnData data) =>
                {
                    // Arrange - Create fissi sheet with text values in various columns
                    using (var package = new ExcelPackage())
                    {
                        var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                        var targetWorksheet = package.Workbook.Worksheets.Add("target");

                        // Add header row
                        fissiWorksheet.Cells[2, 1].Value = "Data";
                        fissiWorksheet.Cells[2, 2].Value = "Partenza";
                        fissiWorksheet.Cells[2, 3].Value = "Assistito";
                        fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                        // Add data row with text values (including in columns 2 and 9)
                        fissiWorksheet.Cells[3, 1].Value = data.DataText;
                        fissiWorksheet.Cells[3, 2].Value = data.PartenzaText;  // Text in time column
                        fissiWorksheet.Cells[3, 3].Value = data.AssistitoText;
                        fissiWorksheet.Cells[3, 9].Value = data.ArrivoText;    // Text in time column

                        var fissiSheet = new Sheet(fissiWorksheet);
                        var targetSheet = new Sheet(targetWorksheet);

                        // Act
                        _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                        // Assert - Verify text values are preserved
                        var dataValue = targetWorksheet.Cells[1, 1].Value?.ToString();
                        var partenzaValue = targetWorksheet.Cells[1, 2].Value?.ToString();
                        var assistitoValue = targetWorksheet.Cells[1, 3].Value?.ToString();
                        var arrivoValue = targetWorksheet.Cells[1, 9].Value?.ToString();

                        return dataValue == data.DataText &&
                               partenzaValue == data.PartenzaText &&
                               assistitoValue == data.AssistitoText &&
                               arrivoValue == data.ArrivoText;
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 2: Preservation - Date Values Preserve Date Format
        /// 
        /// For any fissi sheet data where column 1 contains date values with date format,
        /// the AppendFissiData method SHALL continue to preserve the date format correctly.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.4**
        /// </summary>
        [Test]
        public void Property_Preservation_DateValuesPreserveFormat()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";

                // Add data row with date value and date format
                var sourceCell = fissiWorksheet.Cells[3, 1];
                sourceCell.Value = new DateTime(2024, 1, 15);
                sourceCell.Style.Numberformat.Format = "dd/mm/yyyy";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify date format is preserved
                var targetCell = targetWorksheet.Cells[1, 1];
                
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("dd/mm/yyyy"),
                    "Date format should be preserved for column 1");
                
                // Verify the value is still a date
                Assert.That(targetCell.Value, Is.InstanceOf<DateTime>(),
                    "Date value should be preserved as DateTime");
            }
        }

        /// <summary>
        /// Property 2: Preservation - Font Properties Are Copied Correctly
        /// 
        /// For any fissi sheet data with font properties (bold, italic, size, name, color),
        /// the AppendFissiData method SHALL continue to copy these properties correctly.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_FontPropertiesCopied()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 3].Value = "Assistito";

                // Add data row with font properties in non-time column
                var sourceCell = fissiWorksheet.Cells[3, 3];
                sourceCell.Value = "Rossi Mario";
                sourceCell.Style.Font.Bold = true;
                sourceCell.Style.Font.Italic = true;
                sourceCell.Style.Font.Size = 14;
                sourceCell.Style.Font.Name = "Arial";
                sourceCell.Style.Font.Color.SetColor(System.Drawing.Color.Blue);

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify font properties are copied
                var targetCell = targetWorksheet.Cells[1, 3];
                
                Assert.That(targetCell.Style.Font.Bold, Is.True,
                    "Font bold should be copied");
                Assert.That(targetCell.Style.Font.Italic, Is.True,
                    "Font italic should be copied");
                Assert.That(targetCell.Style.Font.Size, Is.EqualTo(14),
                    "Font size should be copied");
                Assert.That(targetCell.Style.Font.Name, Is.EqualTo("Arial"),
                    "Font name should be copied");
            }
        }

        /// <summary>
        /// Property 2: Preservation - Border Properties Are Copied Correctly
        /// 
        /// For any fissi sheet data with border properties,
        /// the AppendFissiData method SHALL continue to copy these properties correctly.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_BorderPropertiesCopied()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";

                // Add data row with border properties
                var sourceCell = fissiWorksheet.Cells[3, 1];
                sourceCell.Value = "15/01/2024";
                sourceCell.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sourceCell.Style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                sourceCell.Style.Border.Left.Style = ExcelBorderStyle.Dotted;
                sourceCell.Style.Border.Right.Style = ExcelBorderStyle.Dashed;

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify border properties are copied
                var targetCell = targetWorksheet.Cells[1, 1];
                
                Assert.That(targetCell.Style.Border.Top.Style, Is.EqualTo(ExcelBorderStyle.Thin),
                    "Top border should be copied");
                Assert.That(targetCell.Style.Border.Bottom.Style, Is.EqualTo(ExcelBorderStyle.Thick),
                    "Bottom border should be copied");
                Assert.That(targetCell.Style.Border.Left.Style, Is.EqualTo(ExcelBorderStyle.Dotted),
                    "Left border should be copied");
                Assert.That(targetCell.Style.Border.Right.Style, Is.EqualTo(ExcelBorderStyle.Dashed),
                    "Right border should be copied");
            }
        }

        /// <summary>
        /// Property 2: Preservation - Row Structure Maintained
        /// 
        /// For any fissi sheet with multiple rows, the AppendFissiData method SHALL continue
        /// to maintain the correct row structure and column mapping for all columns.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_RowStructureMaintained()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            Prop.ForAll(
                ArbitraryMultipleRows(),
                (MultipleRowsData data) =>
                {
                    // Arrange
                    using (var package = new ExcelPackage())
                    {
                        var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                        var targetWorksheet = package.Workbook.Worksheets.Add("target");

                        // Add header row
                        fissiWorksheet.Cells[2, 1].Value = "Data";
                        fissiWorksheet.Cells[2, 3].Value = "Assistito";
                        fissiWorksheet.Cells[2, 4].Value = "Indirizzo";

                        // Add data rows
                        var rowsList = Microsoft.FSharp.Collections.ListModule.ToArray(data.Rows);
                        for (int i = 0; i < rowsList.Length; i++)
                        {
                            var row = rowsList[i];
                            int fissiRow = 3 + i;
                            fissiWorksheet.Cells[fissiRow, 1].Value = row.Data;
                            fissiWorksheet.Cells[fissiRow, 3].Value = row.Assistito;
                            fissiWorksheet.Cells[fissiRow, 4].Value = row.Indirizzo;
                        }

                        var fissiSheet = new Sheet(fissiWorksheet);
                        var targetSheet = new Sheet(targetWorksheet);

                        // Act
                        _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                        // Assert - Verify all rows maintain correct structure
                        bool allRowsCorrect = true;
                        for (int i = 0; i < rowsList.Length; i++)
                        {
                            var row = rowsList[i];
                            int targetRow = 1 + i;
                            
                            var dataValue = targetWorksheet.Cells[targetRow, 1].Value?.ToString();
                            var assistitoValue = targetWorksheet.Cells[targetRow, 3].Value?.ToString();
                            var indirizzoValue = targetWorksheet.Cells[targetRow, 4].Value?.ToString();

                            if (dataValue != row.Data || 
                                assistitoValue != row.Assistito || 
                                indirizzoValue != row.Indirizzo)
                            {
                                allRowsCorrect = false;
                                break;
                            }
                        }

                        return allRowsCorrect;
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 2: Preservation - Numeric Values (Non-Time) Preserve Format
        /// 
        /// For any fissi sheet data where non-time columns contain numeric values,
        /// the AppendFissiData method SHALL continue to preserve the numeric format and value.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_NumericValuesPreserveFormat()
        {
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 7].Value = "Auto";

                // Add data row with numeric value and format in non-time column
                var sourceCell = fissiWorksheet.Cells[3, 7];
                sourceCell.Value = 123.45;
                sourceCell.Style.Numberformat.Format = "0.00";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify numeric format is preserved
                var targetCell = targetWorksheet.Cells[1, 7];
                
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("0.00"),
                    "Numeric format should be preserved for non-time columns");
                
                Assert.That(targetCell.Value, Is.EqualTo(123.45),
                    "Numeric value should be preserved");
            }
        }

        #region Arbitrary Generators

        /// <summary>
        /// Generator for non-time column data (columns 1, 3, 4, 5, 6, 7, 8)
        /// </summary>
        private static Arbitrary<NonTimeColumnData> ArbitraryNonTimeColumnData()
        {
            var gen = from data in Gen.Elements("15/01/2024", "16/01/2024", "17/01/2024", "18/01/2024")
                      from assistito in Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Giuseppe", "Neri Anna")
                      from indirizzo in Gen.Elements("Via Roma 1", "Via Milano 2", "Via Torino 3", "Via Napoli 4")
                      from destinazione in Gen.Elements("Ospedale", "Clinica", "Casa", "Centro")
                      from note in Gen.Elements("Note 1", "Note 2", "Note 3", "")
                      from auto in Gen.Elements("Auto 1", "Auto 2", "Auto 3", "")
                      from volontario in Gen.Elements("Volontario 1", "Volontario 2", "Volontario 3", "")
                      select new NonTimeColumnData
                      {
                          Data = data,
                          Assistito = assistito,
                          Indirizzo = indirizzo,
                          Destinazione = destinazione,
                          Note = note,
                          Auto = auto,
                          Volontario = volontario
                      };

            return Arb.From(gen);
        }

        /// <summary>
        /// Generator for text column data (including text in time columns 2 and 9)
        /// </summary>
        private static Arbitrary<TextColumnData> ArbitraryTextColumnData()
        {
            var gen = from dataText in Gen.Elements("15/01/2024", "16/01/2024", "17/01/2024")
                      from partenzaText in Gen.Elements("N/A", "TBD", "Pending", "")
                      from assistitoText in Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Giuseppe")
                      from arrivoText in Gen.Elements("N/A", "TBD", "Pending", "")
                      select new TextColumnData
                      {
                          DataText = dataText,
                          PartenzaText = partenzaText,
                          AssistitoText = assistitoText,
                          ArrivoText = arrivoText
                      };

            return Arb.From(gen);
        }

        /// <summary>
        /// Generator for multiple rows (1-5 rows)
        /// </summary>
        private static Arbitrary<MultipleRowsData> ArbitraryMultipleRows()
        {
            var rowGen = from data in Gen.Elements("15/01/2024", "16/01/2024", "17/01/2024")
                         from assistito in Gen.Elements("Rossi Mario", "Bianchi Luigi", "Verdi Giuseppe")
                         from indirizzo in Gen.Elements("Via Roma 1", "Via Milano 2", "Via Torino 3")
                         select new RowData
                         {
                             Data = data,
                             Assistito = assistito,
                             Indirizzo = indirizzo
                         };

            var gen = from count in Gen.Choose(1, 5)
                      from rows in Gen.ListOf(count, rowGen)
                      select new MultipleRowsData { Rows = rows };

            return Arb.From(gen);
        }

        #endregion

        #region Test Data Classes

        public class NonTimeColumnData
        {
            public string Data { get; set; } = null!;
            public string Assistito { get; set; } = null!;
            public string Indirizzo { get; set; } = null!;
            public string Destinazione { get; set; } = null!;
            public string Note { get; set; } = null!;
            public string Auto { get; set; } = null!;
            public string Volontario { get; set; } = null!;
        }

        public class TextColumnData
        {
            public string DataText { get; set; } = null!;
            public string PartenzaText { get; set; } = null!;
            public string AssistitoText { get; set; } = null!;
            public string ArrivoText { get; set; } = null!;
        }

        public class RowData
        {
            public string Data { get; set; } = null!;
            public string Assistito { get; set; } = null!;
            public string Indirizzo { get; set; } = null!;
        }

        public class MultipleRowsData
        {
            public Microsoft.FSharp.Collections.FSharpList<RowData> Rows { get; set; } = null!;
        }

        #endregion
    }
}
