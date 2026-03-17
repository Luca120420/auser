using System;
using System.Collections.Generic;
using System.Linq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for the VLOOKUP formula feature in WriteDataRowsEnhanced.
    /// Validates: Requirements 1.1, 1.2, 1.4, 2.1, 2.2, 2.4, 3.1, 3.2, 4.1, 4.2, 4.3
    /// </summary>
    [TestFixture]
    public class VlookupIndirizzoNotePropertyTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Genera una singola EnhancedTransformedRow con valori arbitrari.
        /// </summary>
        private static Gen<EnhancedTransformedRow> RowGen()
        {
            var strGen = Gen.Elements("", "A", "Rossi Mario", "Via Roma 1", "Ospedale");
            return strGen.SelectMany(partenza =>
                strGen.SelectMany(assistito =>
                strGen.SelectMany(indirizzo =>
                strGen.SelectMany(destinazione =>
                strGen.SelectMany(note =>
                strGen.SelectMany(auto =>
                strGen.SelectMany(volontario =>
                strGen.SelectMany(arrivo =>
                strGen.SelectMany(avv =>
                strGen.Select(indirizzoGasnet => new EnhancedTransformedRow
                {
                    Data = "15/01/2024",
                    Partenza = partenza,
                    Assistito = assistito,
                    Indirizzo = indirizzo,
                    Destinazione = destinazione,
                    Note = note,
                    Auto = auto,
                    Volontario = volontario,
                    Arrivo = arrivo,
                    Avv = avv,
                    IndirizzoGasnet = indirizzoGasnet,
                    NoteGasnet = ""
                }))))))))));
        }

        /// <summary>
        /// Genera una List&lt;EnhancedTransformedRow&gt; di lunghezza 0–20.
        /// </summary>
        private static Gen<List<EnhancedTransformedRow>> RowListGen()
        {
            return Gen.Choose(0, 20).SelectMany(count =>
                Gen.Sequence(Enumerable.Repeat(RowGen(), count)).Select(rows => rows.ToList()));
        }

        // Feature: vlookup-indirizzo-note, Property 1: VLOOKUP formulas written for Indirizzo and Note
        /// <summary>
        /// Per qualsiasi lista di righe (0–20) e startRow (2–100),
        /// ogni cella in col 4 e col 6 ha Formula non vuota contenente "VLOOKUP".
        /// Validates: Requirements 1.1, 1.4, 2.1, 2.4
        /// </summary>
        [Test]
        public void Property1_Col4AndCol6_HaveVlookupFormula()
        {
            var arb = Arb.From(
                Gen.Zip(RowListGen(), Gen.Choose(2, 100))
            );

            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(arb, tuple =>
            {
                var (rows, startRow) = tuple;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    var sheet = new Sheet(worksheet);

                    _excelManager.WriteDataRowsEnhanced(sheet, rows, startRow);

                    for (int i = 0; i < rows.Count; i++)
                    {
                        int excelRow = startRow + i;

                        var formula4 = worksheet.Cells[excelRow, 4].Formula;
                        if (string.IsNullOrEmpty(formula4) || !formula4.Contains("VLOOKUP"))
                            return false;

                        var formula6 = worksheet.Cells[excelRow, 6].Formula;
                        if (string.IsNullOrEmpty(formula6) || !formula6.Contains("VLOOKUP"))
                            return false;
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: vlookup-indirizzo-note, Property 2: Row number accuracy in formula
        /// <summary>
        /// Per qualsiasi lista di N righe e startRow, la formula nella riga i-esima (0-indexed) deve essere:
        /// Col 4: VLOOKUP(C{startRow+i},assistiti!$A:$C,2,0)
        /// Col 6: VLOOKUP(C{startRow+i},assistiti!$A:$C,3,0)
        /// Validates: Requirements 1.2, 2.2, 4.1, 4.2, 4.3
        /// </summary>
        [Test]
        public void Property2_FormulaContainsCorrectRowNumber()
        {
            var arb = Arb.From(
                Gen.Zip(RowListGen(), Gen.Choose(2, 100))
            );

            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(arb, tuple =>
            {
                var (rows, startRow) = tuple;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    var sheet = new Sheet(worksheet);

                    _excelManager.WriteDataRowsEnhanced(sheet, rows, startRow);

                    for (int i = 0; i < rows.Count; i++)
                    {
                        int excelRow = startRow + i;
                        string expectedFormula4 = $"VLOOKUP(C{excelRow},assistiti!A:C,2,FALSE)";
                        string expectedFormula6 = $"VLOOKUP(C{excelRow},assistiti!A:C,3,FALSE)";

                        if (worksheet.Cells[excelRow, 4].Formula != expectedFormula4)
                            return false;

                        if (worksheet.Cells[excelRow, 6].Formula != expectedFormula6)
                            return false;
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: vlookup-indirizzo-note, Property 3: Other columns unaffected
        /// <summary>
        /// Per qualsiasi lista di righe, le colonne 1,2,3,5,7,8,9,10,11,12 hanno Formula vuota.
        /// Validates: Requirements 3.1, 3.2
        /// </summary>
        [Test]
        public void Property3_OtherColumns_HaveEmptyFormula()
        {
            var arb = Arb.From(
                Gen.Zip(RowListGen(), Gen.Choose(2, 100))
            );

            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            int[] nonFormulaCols = { 1, 2, 3, 5, 7, 8, 9, 10, 11, 12 };

            Prop.ForAll(arb, tuple =>
            {
                var (rows, startRow) = tuple;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    var sheet = new Sheet(worksheet);

                    _excelManager.WriteDataRowsEnhanced(sheet, rows, startRow);

                    for (int i = 0; i < rows.Count; i++)
                    {
                        int excelRow = startRow + i;
                        foreach (var col in nonFormulaCols)
                        {
                            var formula = worksheet.Cells[excelRow, col].Formula;
                            if (!string.IsNullOrEmpty(formula))
                                return false;
                        }
                    }

                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }
    }
}
