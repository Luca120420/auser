using System;
using System.Collections.Generic;
using System.Linq;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for LookupService class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 8.2, 8.3
    /// </summary>
    [TestFixture]
    public class LookupServicePropertyTests
    {
        private LookupService _lookupService = null!;

        [SetUp]
        public void Setup()
        {
            _lookupService = new LookupService();
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Generator for assistito data with Italian characters
        /// </summary>
        public class AssistitoData
        {
            public string Nome { get; set; } = "";
            public string Indirizzo { get; set; } = "";
            public string Note { get; set; } = "";
        }

        /// <summary>
        /// Arbitrary generator for assistito data with Italian characters
        /// </summary>
        public static Arbitrary<AssistitoData> ArbitraryAssistitoData()
        {
            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };
            var normalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '-.,";

            // Generator for strings that may contain Italian characters
            var italianStringGen = from length in Gen.Choose(1, 50)
                                   from useItalian in Arb.Generate<bool>()
                                   from chars in Gen.ArrayOf(length, useItalian && length > 0
                                       ? Gen.Elements(italianChars.Concat(normalChars.ToCharArray()).ToArray())
                                       : Gen.Elements(normalChars.ToCharArray()))
                                   select new string(chars).Trim();

            // Generator for assistito names (non-empty)
            var nomeGen = from nome in italianStringGen
                         where !string.IsNullOrWhiteSpace(nome)
                         select nome;

            var assistitoGen = from nome in nomeGen
                              from indirizzo in italianStringGen
                              from note in italianStringGen
                              select new AssistitoData
                              {
                                  Nome = nome,
                                  Indirizzo = indirizzo,
                                  Note = note
                              };

            return Arb.From(assistitoGen);
        }

        /// <summary>
        /// Helper method to create a test assistiti sheet with data
        /// </summary>
        private Sheet CreateAssistitiSheet(List<AssistitoData> assistitiData)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");

            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";

            // Write data rows
            for (int i = 0; i < assistitiData.Count; i++)
            {
                var assistito = assistitiData[i];
                worksheet.Cells[i + 2, 1].Value = assistito.Nome;
                worksheet.Cells[i + 2, 2].Value = assistito.Indirizzo;
                worksheet.Cells[i + 2, 3].Value = assistito.Note;
            }

            return new Sheet(worksheet);
        }

        /// <summary>
        /// Helper method to create an empty fissi sheet
        /// </summary>
        private Sheet CreateEmptyFissiSheet()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("fissi");
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Avv";
            return new Sheet(worksheet);
        }

        /// <summary>
        /// Generator for fissi data with Italian characters
        /// </summary>
        public class FissiData
        {
            public string Nome { get; set; } = "";
            public string Avv { get; set; } = "";
        }

        /// <summary>
        /// Arbitrary generator for fissi data with Italian characters
        /// </summary>
        public static Arbitrary<FissiData> ArbitraryFissiData()
        {
            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };
            var normalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '-.,";

            // Generator for strings that may contain Italian characters
            var italianStringGen = from length in Gen.Choose(1, 50)
                                   from useItalian in Arb.Generate<bool>()
                                   from chars in Gen.ArrayOf(length, useItalian && length > 0
                                       ? Gen.Elements(italianChars.Concat(normalChars.ToCharArray()).ToArray())
                                       : Gen.Elements(normalChars.ToCharArray()))
                                   select new string(chars).Trim();

            // Generator for fissi names (non-empty)
            var nomeGen = from nome in italianStringGen
                         where !string.IsNullOrWhiteSpace(nome)
                         select nome;

            var fissiGen = from nome in nomeGen
                          from avv in italianStringGen
                          select new FissiData
                          {
                              Nome = nome,
                              Avv = avv
                          };

            return Arb.From(fissiGen);
        }

        /// <summary>
        /// Helper method to create a test fissi sheet with data
        /// </summary>
        private Sheet CreateFissiSheet(List<FissiData> fissiData)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("fissi");

            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Avv";

            // Write data rows
            for (int i = 0; i < fissiData.Count; i++)
            {
                var fissi = fissiData[i];
                worksheet.Cells[i + 2, 1].Value = fissi.Nome;
                worksheet.Cells[i + 2, 2].Value = fissi.Avv;
            }

            return new Sheet(worksheet);
        }

        // Feature: excel-output-enhancement, Property 4: Assistiti Lookup Correctness
        /// <summary>
        /// Property 4: Assistiti Lookup Correctness
        /// For any assistito name that exists in the "assistiti" reference sheet,
        /// the corresponding row's "Indirizzo" and "Note" values SHALL match the values
        /// from the reference sheet for that assistito.
        /// **Validates: Requirements 5.1, 6.1**
        /// </summary>
        [Test]
        public void Property_AssistitiLookupCorrectness()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryAssistitoData().Generator).Select(list => list.ToList())),
                (List<AssistitoData> assistitiData) =>
                {
                    try
                    {
                        // Arrange - Create assistiti sheet with test data
                        var assistitiSheet = CreateAssistitiSheet(assistitiData);
                        var fissiSheet = CreateEmptyFissiSheet();
                        
                        _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

                        // Act & Assert - For each UNIQUE assistito (first occurrence only),
                        // verify lookup returns correct values
                        // This matches the LookupService behavior of using the first match for duplicates
                        var uniqueAssistitiData = new Dictionary<string, AssistitoData>();
                        foreach (var assistito in assistitiData)
                        {
                            if (!uniqueAssistitiData.ContainsKey(assistito.Nome))
                            {
                                uniqueAssistitiData[assistito.Nome] = assistito;
                            }
                        }

                        foreach (var kvp in uniqueAssistitiData)
                        {
                            var assistito = kvp.Value;
                            var lookedUpIndirizzo = _lookupService.LookupInAssistiti(assistito.Nome, "Indirizzo");
                            var lookedUpNote = _lookupService.LookupInAssistiti(assistito.Nome, "Note");

                            // Normalize empty strings (treat null and empty as equivalent)
                            var expectedIndirizzo = assistito.Indirizzo ?? "";
                            var expectedNote = assistito.Note ?? "";

                            if (lookedUpIndirizzo != expectedIndirizzo)
                            {
                                return false.Label($"Indirizzo mismatch for '{assistito.Nome}': expected '{expectedIndirizzo}', got '{lookedUpIndirizzo}'");
                            }

                            if (lookedUpNote != expectedNote)
                            {
                                return false.Label($"Note mismatch for '{assistito.Nome}': expected '{expectedNote}', got '{lookedUpNote}'");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Lookup failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: excel-output-enhancement, Property 5: Missing Lookup Key Handling
        /// <summary>
        /// Property 5: Missing Lookup Key Handling
        /// For any lookup operation (assistiti or fissi) where the lookup key is not found
        /// in the reference sheet, the corresponding cell SHALL be empty.
        /// **Validates: Requirements 5.2, 6.2, 8.3**
        /// </summary>
        [Test]
        public void Property_MissingLookupKeyHandling()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryAssistitoData().Generator).Select(list => list.ToList())),
                Arb.Default.NonEmptyString(),
                (List<AssistitoData> assistitiData, NonEmptyString missingKeyWrapper) =>
                {
                    try
                    {
                        // Arrange - Create reference sheets with known data
                        var assistitiSheet = CreateAssistitiSheet(assistitiData);
                        var fissiSheet = CreateEmptyFissiSheet();
                        
                        _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

                        // Generate a lookup key that is guaranteed not to exist in the reference sheets
                        var existingKeys = assistitiData.Select(a => a.Nome).ToHashSet();
                        var missingKey = missingKeyWrapper.Get;
                        
                        // Ensure the key doesn't exist by appending a unique suffix if needed
                        while (existingKeys.Contains(missingKey))
                        {
                            missingKey = missingKey + "_MISSING_" + Guid.NewGuid().ToString();
                        }

                        // Act - Perform lookups with missing key
                        var indirizzoResult = _lookupService.LookupInAssistiti(missingKey, "Indirizzo");
                        var noteResult = _lookupService.LookupInAssistiti(missingKey, "Note");
                        var avvResult = _lookupService.LookupInFissi(missingKey, "Avv");

                        // Assert - All missing lookups should return empty string
                        if (indirizzoResult != "")
                        {
                            return false.Label($"Assistiti lookup for missing key '{missingKey}' (Indirizzo) should return empty string, got '{indirizzoResult}'");
                        }

                        if (noteResult != "")
                        {
                            return false.Label($"Assistiti lookup for missing key '{missingKey}' (Note) should return empty string, got '{noteResult}'");
                        }

                        if (avvResult != "")
                        {
                            return false.Label($"Fissi lookup for missing key '{missingKey}' (Avv) should return empty string, got '{avvResult}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Missing lookup key handling failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: excel-output-enhancement, Property 7: Fissi Lookup Correctness
        /// <summary>
        /// Property 7: Fissi Lookup Correctness
        /// For any lookup key that exists in the "fissi" reference sheet,
        /// the corresponding row's "Avv" value SHALL match the value
        /// from the reference sheet for that key.
        /// **Validates: Requirements 8.2**
        /// </summary>
        [Test]
        public void Property_FissiLookupCorrectness()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryFissiData().Generator).Select(list => list.ToList())),
                (List<FissiData> fissiData) =>
                {
                    try
                    {
                        // Arrange - Create fissi sheet with test data
                        var assistitiSheet = CreateAssistitiSheet(new List<AssistitoData>());
                        var fissiSheet = CreateFissiSheet(fissiData);
                        
                        _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

                        // Act & Assert - For each UNIQUE fissi entry (first occurrence only), 
                        // verify lookup returns correct Avv value
                        // This matches the LookupService behavior of using the first match for duplicates
                        var uniqueFissiData = new Dictionary<string, FissiData>();
                        foreach (var fissi in fissiData)
                        {
                            if (!uniqueFissiData.ContainsKey(fissi.Nome))
                            {
                                uniqueFissiData[fissi.Nome] = fissi;
                            }
                        }

                        foreach (var kvp in uniqueFissiData)
                        {
                            var fissi = kvp.Value;
                            var lookedUpAvv = _lookupService.LookupInFissi(fissi.Nome, "Avv");

                            // Normalize empty strings (treat null and empty as equivalent)
                            var expectedAvv = fissi.Avv ?? "";

                            if (lookedUpAvv != expectedAvv)
                            {
                                return false.Label($"Avv mismatch for '{fissi.Nome}': expected '{expectedAvv}', got '{lookedUpAvv}'");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Fissi lookup failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }
    }
}
