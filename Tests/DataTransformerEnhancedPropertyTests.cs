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
    /// Property-based tests for DataTransformer enhanced transformation using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 7.2
    /// </summary>
    [TestFixture]
    public class DataTransformerEnhancedPropertyTests
    {
        private DataTransformer _dataTransformer = null!;
        private LookupService _lookupService = null!;
        private TransformationRulesEngine _rulesEngine = null!;

        [SetUp]
        public void Setup()
        {
            _rulesEngine = new TransformationRulesEngine();
            _dataTransformer = new DataTransformer(_rulesEngine);
            _lookupService = new LookupService();
            
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            // Load empty reference sheets for lookup service
            var emptyAssistitiSheet = CreateEmptyAssistitiSheet();
            var emptyFissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(emptyAssistitiSheet, emptyFissiSheet);
        }

        /// <summary>
        /// Helper method to create an empty assistiti sheet
        /// </summary>
        private Sheet CreateEmptyAssistitiSheet()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";
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
        /// Generator for ServiceAppointment with random Note field values
        /// </summary>
        public class AppointmentWithNote
        {
            public ServiceAppointment Appointment { get; set; } = null!;
            public string ExpectedNoteGasnet { get; set; } = "";
        }

        /// <summary>
        /// Arbitrary generator for ServiceAppointment with Italian characters in Note field
        /// </summary>
        public static Arbitrary<AppointmentWithNote> ArbitraryAppointmentWithNote()
        {
            var italianChars = new[] { 'à', 'è', 'é', 'ì', 'ò', 'ù', 'À', 'È', 'É', 'Ì', 'Ò', 'Ù' };
            var normalChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789 '-.,!?";

            // Generator for strings that may contain Italian characters
            var italianStringGen = from length in Gen.Choose(0, 100)
                                   from useItalian in Arb.Generate<bool>()
                                   from chars in Gen.ArrayOf(length, useItalian && length > 0
                                       ? Gen.Elements(italianChars.Concat(normalChars.ToCharArray()).ToArray())
                                       : Gen.Elements(normalChars.ToCharArray()))
                                   select new string(chars).Trim();

            // Generator for non-empty names
            var nameGen = from length in Gen.Choose(1, 20)
                         from chars in Gen.ArrayOf(length, Gen.Elements(normalChars.ToCharArray()))
                         let name = new string(chars).Trim()
                         where !string.IsNullOrWhiteSpace(name)
                         select name;

            // Generator for dates in dd/MM/yyyy format
            var dateGen = from year in Gen.Choose(2024, 2026)
                         from month in Gen.Choose(1, 12)
                         from day in Gen.Choose(1, 28) // Use 28 to avoid invalid dates
                         select $"{day:D2}/{month:D2}/{year}";

            // Generator for times in HH:mm format
            var timeGen = from hour in Gen.Choose(0, 23)
                         from minute in Gen.Choose(0, 59)
                         select $"{hour:D2}:{minute:D2}";

            var appointmentGen = from cognome in nameGen
                                from nome in nameGen
                                from note in italianStringGen
                                from date in dateGen
                                from time in timeGen
                                select new AppointmentWithNote
                                {
                                    Appointment = new ServiceAppointment
                                    {
                                        DataServizio = date,
                                        OraInizioServizio = time,
                                        CognomeAssistito = cognome,
                                        NomeAssistito = nome,
                                        NoteERichieste = note,
                                        Attivita = "Test Activity",
                                        DescrizioneStatoServizio = "ATTIVO",
                                        IndirizzoPartenza = "Via Test 123",
                                        ComunePartenza = "Milano",
                                        IndirizzoDestinazione = "Via Dest 456",
                                        ComuneDestinazione = "Roma",
                                        CausaleDestinazione = "Visita medica"
                                    },
                                    ExpectedNoteGasnet = note ?? ""
                                };

            return Arb.From(appointmentGen);
        }

        // Feature: excel-output-enhancement, Property 6: CSV Note Mapping
        /// <summary>
        /// Property 6: CSV Note Mapping
        /// For any CSV row with a Note field value, the corresponding output row's
        /// "Note Gasnet" column SHALL contain that exact value.
        /// **Validates: Requirements 7.2**
        /// </summary>
        [Test]
        public void Property_CSVNoteMappingToNoteGasnet()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(Gen.NonEmptyListOf(ArbitraryAppointmentWithNote().Generator).Select(list => list.ToList())),
                (List<AppointmentWithNote> appointmentsWithNotes) =>
                {
                    try
                    {
                        // Arrange - Extract appointments
                        var appointments = appointmentsWithNotes.Select(a => a.Appointment).ToList();

                        // Act - Transform appointments
                        var result = _dataTransformer.TransformEnhanced(appointments, _lookupService);

                        // Assert - Verify each row's NoteGasnet matches the CSV Note field
                        for (int i = 0; i < appointmentsWithNotes.Count; i++)
                        {
                            var expected = appointmentsWithNotes[i].ExpectedNoteGasnet;
                            var actual = result.Rows[i].NoteGasnet;

                            if (actual != expected)
                            {
                                return false.Label($"Row {i}: NoteGasnet mismatch. Expected '{expected}', got '{actual}'");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"CSV Note mapping failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }
    }
}
