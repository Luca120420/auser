using System;
using System.Collections.Generic;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the VLOOKUP formula feature in WriteDataRowsEnhanced.
    /// Verifies that col 4 (Indirizzo) and col 6 (Note) receive VLOOKUP formulas.
    /// </summary>
    [TestFixture]
    public class VlookupIndirizzoNoteTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        // Task 2.1 — singola riga con startRow=2: formula esatta in col 4 e col 6
        [Test]
        public void WriteDataRowsEnhanced_SingleRow_StartRow2_Col4HasExactVlookupFormula()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "15/01/2024",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario",
                        Indirizzo = "Via Roma 1",
                        Destinazione = "Ospedale",
                        Note = "Nota test",
                        Auto = "Auto1",
                        Volontario = "Vol1",
                        Arrivo = "10:00",
                        Avv = "",
                        IndirizzoGasnet = "",
                        NoteGasnet = ""
                    }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 2);

                Assert.That(worksheet.Cells[2, 4].Formula,
                    Is.EqualTo("VLOOKUP(C2,assistiti!A:C,2,FALSE)"),
                    "Col 4 deve avere la formula VLOOKUP esatta per la riga 2");
            }
        }

        [Test]
        public void WriteDataRowsEnhanced_SingleRow_StartRow2_Col6HasExactVlookupFormula()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "15/01/2024",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario",
                        Indirizzo = "Via Roma 1",
                        Destinazione = "Ospedale",
                        Note = "Nota test",
                        Auto = "Auto1",
                        Volontario = "Vol1",
                        Arrivo = "10:00",
                        Avv = "",
                        IndirizzoGasnet = "",
                        NoteGasnet = ""
                    }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 2);

                Assert.That(worksheet.Cells[2, 6].Formula,
                    Is.EqualTo("VLOOKUP(C2,assistiti!A:C,3,FALSE)"),
                    "Col 6 deve avere la formula VLOOKUP esatta per la riga 2");
            }
        }

        // Task 2.2 — colonne non interessate hanno Formula vuota
        [Test]
        public void WriteDataRowsEnhanced_SingleRow_StartRow2_NonFormulaCols_HaveEmptyFormula()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                var rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "15/01/2024",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario",
                        Indirizzo = "Via Roma 1",
                        Destinazione = "Ospedale",
                        Note = "Nota test",
                        Auto = "Auto1",
                        Volontario = "Vol1",
                        Arrivo = "10:00",
                        Avv = "Avv1",
                        IndirizzoGasnet = "Via Gasnet 1",
                        NoteGasnet = "NoteGasnet1"
                    }
                };

                _excelManager.WriteDataRowsEnhanced(sheet, rows, 2);

                int[] nonFormulaCols = { 1, 2, 3, 5, 7, 8, 9, 10, 11, 12 };
                foreach (var col in nonFormulaCols)
                {
                    var formula = worksheet.Cells[2, col].Formula;
                    Assert.That(string.IsNullOrEmpty(formula), Is.True,
                        $"Col {col} non deve avere formula, ma ha: '{formula}'");
                }
            }
        }

        // Task 2.3 — lista vuota: nessuna eccezione
        [Test]
        public void WriteDataRowsEnhanced_EmptyList_DoesNotThrow()
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Test");
                var sheet = new Sheet(worksheet);

                Assert.DoesNotThrow(() =>
                    _excelManager.WriteDataRowsEnhanced(sheet, new List<EnhancedTransformedRow>(), 2));
            }
        }
    }
}
