using System;
using System.Collections.Generic;
using NUnit.Framework;
using Moq;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for DataTransformer enhanced transformation.
    /// Tests specific examples and edge cases.
    /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 7.2, 8.2, 8.3
    /// </summary>
    [TestFixture]
    public class DataTransformerEnhancedTests
    {
        private DataTransformer _dataTransformer = null!;
        private Mock<ILookupService> _mockLookupService = null!;
        private TransformationRulesEngine _rulesEngine = null!;

        [SetUp]
        public void Setup()
        {
            _rulesEngine = new TransformationRulesEngine();
            _dataTransformer = new DataTransformer(_rulesEngine);
            _mockLookupService = new Mock<ILookupService>();
            
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        /// <summary>
        /// Test transformation with valid assistiti and fissi data.
        /// Verifies that lookups are performed correctly and all columns are populated.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithValidLookupData_PopulatesAllColumns()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "15/01/2024",
                OraInizioServizio = "09:30",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                NoteERichieste = "Portare documenti",
                Attivita = "Accompagnamento",
                DescrizioneStatoServizio = "ATTIVO",
                IndirizzoPartenza = "Via Roma 10",
                ComunePartenza = "Milano",
                IndirizzoDestinazione = "Via Verdi 20",
                ComuneDestinazione = "Roma",
                CausaleDestinazione = "Visita medica"
            };

            // Setup mock lookups
            _mockLookupService.Setup(x => x.LookupInAssistiti("Rossi Mario", "Indirizzo"))
                .Returns("Via Garibaldi 5");
            _mockLookupService.Setup(x => x.LookupInAssistiti("Rossi Mario", "Note"))
                .Returns("Cliente abituale");
            _mockLookupService.Setup(x => x.LookupInFissi("Rossi Mario", "Avv"))
                .Returns("SI");

            // Act
            var result = _dataTransformer.TransformEnhanced(new List<ServiceAppointment> { appointment }, _mockLookupService.Object);

            // Assert
            Assert.That(result.Rows, Has.Count.EqualTo(1));
            var row = result.Rows[0];

            Assert.That(row.Data, Is.EqualTo("15/01/2024"));
            Assert.That(row.Partenza, Is.EqualTo("09:30"));
            Assert.That(row.Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(row.Indirizzo, Is.EqualTo("Via Garibaldi 5"));
            Assert.That(row.Destinazione, Is.EqualTo("Roma Via Verdi 20 Visita medica"));
            Assert.That(row.Note, Is.EqualTo("Cliente abituale"));
            Assert.That(row.Auto, Is.Empty);
            Assert.That(row.Volontario, Is.Empty);
            Assert.That(row.Arrivo, Is.Empty);
            Assert.That(row.Avv, Is.EqualTo("SI"));
            Assert.That(row.IndirizzoGasnet, Is.EqualTo("Milano Via Roma 10"));
            Assert.That(row.NoteGasnet, Is.EqualTo("Portare documenti"));

            // Verify lookups were called
            _mockLookupService.Verify(x => x.LookupInAssistiti("Rossi Mario", "Indirizzo"), Times.Once);
            _mockLookupService.Verify(x => x.LookupInAssistiti("Rossi Mario", "Note"), Times.Once);
            _mockLookupService.Verify(x => x.LookupInFissi("Rossi Mario", "Avv"), Times.Once);
        }

        /// <summary>
        /// Test transformation with missing lookup keys.
        /// Verifies that empty strings are returned for missing lookups.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithMissingLookupKeys_ReturnsEmptyStrings()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "16/01/2024",
                OraInizioServizio = "10:00",
                CognomeAssistito = "Bianchi",
                NomeAssistito = "Luigi",
                NoteERichieste = "Nessuna nota",
                Attivita = "Trasporto",
                DescrizioneStatoServizio = "ATTIVO",
                IndirizzoPartenza = "Via Milano 15",
                ComunePartenza = "Torino",
                IndirizzoDestinazione = "Via Napoli 30",
                ComuneDestinazione = "Genova",
                CausaleDestinazione = "Controllo"
            };

            // Setup mock lookups to return empty strings (simulating missing keys)
            _mockLookupService.Setup(x => x.LookupInAssistiti("Bianchi Luigi", "Indirizzo"))
                .Returns("");
            _mockLookupService.Setup(x => x.LookupInAssistiti("Bianchi Luigi", "Note"))
                .Returns("");
            _mockLookupService.Setup(x => x.LookupInFissi("Bianchi Luigi", "Avv"))
                .Returns("");

            // Act
            var result = _dataTransformer.TransformEnhanced(new List<ServiceAppointment> { appointment }, _mockLookupService.Object);

            // Assert
            Assert.That(result.Rows, Has.Count.EqualTo(1));
            var row = result.Rows[0];

            Assert.That(row.Indirizzo, Is.Empty, "Indirizzo should be empty for missing lookup key");
            Assert.That(row.Note, Is.Empty, "Note should be empty for missing lookup key");
            Assert.That(row.Avv, Is.Empty, "Avv should be empty for missing lookup key");
            Assert.That(row.NoteGasnet, Is.EqualTo("Nessuna nota"), "NoteGasnet should contain CSV note");
        }

        /// <summary>
        /// Test transformation with empty CSV Note field.
        /// Verifies that NoteGasnet is empty when CSV Note is null or empty.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithEmptyCSVNote_ReturnsEmptyNoteGasnet()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "17/01/2024",
                OraInizioServizio = "11:00",
                CognomeAssistito = "Verdi",
                NomeAssistito = "Giuseppe",
                NoteERichieste = null, // Empty CSV note
                Attivita = "Accompagnamento",
                DescrizioneStatoServizio = "ATTIVO",
                IndirizzoPartenza = "Via Firenze 25",
                ComunePartenza = "Bologna",
                IndirizzoDestinazione = "Via Venezia 40",
                ComuneDestinazione = "Padova",
                CausaleDestinazione = "Terapia"
            };

            // Setup mock lookups
            _mockLookupService.Setup(x => x.LookupInAssistiti(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");
            _mockLookupService.Setup(x => x.LookupInFissi(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");

            // Act
            var result = _dataTransformer.TransformEnhanced(new List<ServiceAppointment> { appointment }, _mockLookupService.Object);

            // Assert
            Assert.That(result.Rows, Has.Count.EqualTo(1));
            var row = result.Rows[0];

            Assert.That(row.NoteGasnet, Is.Empty, "NoteGasnet should be empty when CSV Note is null");
        }

        /// <summary>
        /// Test transformation with Italian special characters.
        /// Verifies that Italian characters (à, è, é, ì, ò, ù) are handled correctly.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithItalianSpecialCharacters_HandlesCorrectly()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "18/01/2024",
                OraInizioServizio = "14:30",
                CognomeAssistito = "D'Àngelo",
                NomeAssistito = "Nicolò",
                NoteERichieste = "Portare caffè e tè",
                Attivita = "Accompagnamento",
                DescrizioneStatoServizio = "ATTIVO",
                IndirizzoPartenza = "Via dell'Università 5",
                ComunePartenza = "Città di Castello",
                IndirizzoDestinazione = "Via Perché 10",
                ComuneDestinazione = "Così",
                CausaleDestinazione = "Visita più importante"
            };

            // Setup mock lookups with Italian characters
            _mockLookupService.Setup(x => x.LookupInAssistiti("D'Àngelo Nicolò", "Indirizzo"))
                .Returns("Via dell'Università 15");
            _mockLookupService.Setup(x => x.LookupInAssistiti("D'Àngelo Nicolò", "Note"))
                .Returns("Cliente già conosciuto");
            _mockLookupService.Setup(x => x.LookupInFissi("D'Àngelo Nicolò", "Avv"))
                .Returns("Sì");

            // Act
            var result = _dataTransformer.TransformEnhanced(new List<ServiceAppointment> { appointment }, _mockLookupService.Object);

            // Assert
            Assert.That(result.Rows, Has.Count.EqualTo(1));
            var row = result.Rows[0];

            Assert.That(row.Assistito, Is.EqualTo("D'Àngelo Nicolò"));
            Assert.That(row.Indirizzo, Is.EqualTo("Via dell'Università 15"));
            Assert.That(row.Note, Is.EqualTo("Cliente già conosciuto"));
            Assert.That(row.Avv, Is.EqualTo("Sì"));
            Assert.That(row.NoteGasnet, Is.EqualTo("Portare caffè e tè"));
            Assert.That(row.Destinazione, Is.EqualTo("Così Via Perché 10 Visita più importante"));
        }

        /// <summary>
        /// Test that cancelled appointments are filtered out.
        /// Verifies that appointments with "ANNULLATO" status are not included.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithAnnullatoStatus_FiltersOutRow()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "19/01/2024",
                    OraInizioServizio = "09:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    DescrizioneStatoServizio = "ATTIVO",
                    IndirizzoPartenza = "Via Roma 10",
                    ComunePartenza = "Milano"
                },
                new ServiceAppointment
                {
                    DataServizio = "19/01/2024",
                    OraInizioServizio = "10:00",
                    CognomeAssistito = "Bianchi",
                    NomeAssistito = "Luigi",
                    DescrizioneStatoServizio = "ANNULLATO", // Should be filtered out
                    IndirizzoPartenza = "Via Milano 15",
                    ComunePartenza = "Torino"
                },
                new ServiceAppointment
                {
                    DataServizio = "19/01/2024",
                    OraInizioServizio = "11:00",
                    CognomeAssistito = "Verdi",
                    NomeAssistito = "Giuseppe",
                    DescrizioneStatoServizio = "ATTIVO",
                    IndirizzoPartenza = "Via Firenze 25",
                    ComunePartenza = "Bologna"
                }
            };

            // Setup mock lookups
            _mockLookupService.Setup(x => x.LookupInAssistiti(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");
            _mockLookupService.Setup(x => x.LookupInFissi(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");

            // Act
            var result = _dataTransformer.TransformEnhanced(appointments, _mockLookupService.Object);

            // Assert
            Assert.That(result.Rows, Has.Count.EqualTo(2), "Should only include non-cancelled appointments");
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(result.Rows[1].Assistito, Is.EqualTo("Verdi Giuseppe"));
        }

        /// <summary>
        /// Test that yellow highlight rows are tracked correctly.
        /// Verifies that rows with "Accompag. con macchina attrezzata" are marked for highlighting.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithMacchinaAttrezzata_TracksYellowHighlight()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "20/01/2024",
                    OraInizioServizio = "09:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    Attivita = "Accompagnamento normale",
                    DescrizioneStatoServizio = "ATTIVO",
                    IndirizzoPartenza = "Via Roma 10",
                    ComunePartenza = "Milano"
                },
                new ServiceAppointment
                {
                    DataServizio = "20/01/2024",
                    OraInizioServizio = "10:00",
                    CognomeAssistito = "Bianchi",
                    NomeAssistito = "Luigi",
                    Attivita = "Accompag. con macchina attrezzata", // Should be highlighted
                    DescrizioneStatoServizio = "ATTIVO",
                    IndirizzoPartenza = "Via Milano 15",
                    ComunePartenza = "Torino"
                }
            };

            // Setup mock lookups
            _mockLookupService.Setup(x => x.LookupInAssistiti(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");
            _mockLookupService.Setup(x => x.LookupInFissi(It.IsAny<string>(), It.IsAny<string>()))
                .Returns("");

            // Act
            var result = _dataTransformer.TransformEnhanced(appointments, _mockLookupService.Object);

            // Assert
            Assert.That(result.YellowHighlightRows, Has.Count.EqualTo(1));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(4), "Second row should be at index 4 (row 3 is first data row)");
        }

        /// <summary>
        /// Test that null appointments throw ArgumentNullException.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithNullAppointments_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() =>
                _dataTransformer.TransformEnhanced(null!, _mockLookupService.Object));
        }

        /// <summary>
        /// Test that null lookup service throws ArgumentNullException.
        /// </summary>
        [Test]
        public void TransformEnhanced_WithNullLookupService_ThrowsArgumentNullException()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "21/01/2024",
                    OraInizioServizio = "09:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    DescrizioneStatoServizio = "ATTIVO",
                    IndirizzoPartenza = "Via Roma 10",
                    ComunePartenza = "Milano"
                }
            };

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() =>
                _dataTransformer.TransformEnhanced(appointments, null!));
        }
    }
}
