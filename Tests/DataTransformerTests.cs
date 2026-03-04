using System;
using System.Collections.Generic;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the DataTransformer class.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    [TestFixture]
    public class DataTransformerTests
    {
        private ITransformationRulesEngine _rulesEngine = null!;
        private IDataTransformer _dataTransformer = null!;

        [SetUp]
        public void SetUp()
        {
            _rulesEngine = new TransformationRulesEngine();
            _dataTransformer = new DataTransformer(_rulesEngine);
        }

        [Test]
        public void Constructor_WithNullRulesEngine_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => new DataTransformer(null!));
        }

        [Test]
        public void Transform_WithNullAppointments_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _dataTransformer.Transform(null!));
        }

        [Test]
        public void Transform_WithEmptyList_ReturnsEmptyResult()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>();

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows, Is.Not.Null);
            Assert.That(result.YellowHighlightRows, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(0));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(0));
        }

        [Test]
        public void Transform_WithSingleAppointment_ReturnsTransformedRow()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    IndirizzoPartenza = "Via Roma 1",
                    ComunePartenza = "Milano",
                    DescrizionePuntoPartenza = "Casa",
                    IndirizzoDestinazione = "Via Verdi 10",
                    ComuneDestinazione = "Milano",
                    CausaleDestinazione = "Ospedale",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    NoteERichieste = "Nessuna nota"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            
            var row = result.Rows[0];
            Assert.That(row.DataServizio, Is.EqualTo("2024-01-15"));
            Assert.That(row.OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(row.Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(row.CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(row.NomeAssistito, Is.EqualTo("Mario"));
            Assert.That(row.Indirizzo, Is.EqualTo("Milano Via Verdi 10"));
            Assert.That(row.OraInizioServizioCopy, Is.EqualTo("09:00"));
            Assert.That(row.Partenza, Is.EqualTo(string.Empty));
            Assert.That(row.NoteERichieste, Is.EqualTo("Nessuna nota"));
        }

        [Test]
        public void Transform_WithYellowHighlightAppointment_MarksRowForHighlighting()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Accompag. con macchina attrezzata",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "Via Verdi 10"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(1));
        }

        [Test]
        public void Transform_WithAnnullatoAppointment_FiltersOutRow()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "ANNULLATO",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "Via Verdi 10"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(0));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(0));
        }

        [Test]
        public void Transform_WithMultipleAppointments_TransformsAllNonCancelled()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "Via Verdi 10"
                },
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "10:00",
                    Attivita = "Accompag. con macchina attrezzata",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Bianchi",
                    NomeAssistito = "Luigi",
                    ComuneDestinazione = "Roma",
                    IndirizzoDestinazione = "Via Garibaldi 5"
                },
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "11:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "ANNULLATO",
                    CognomeAssistito = "Verdi",
                    NomeAssistito = "Giuseppe",
                    ComuneDestinazione = "Napoli",
                    IndirizzoDestinazione = "Via Dante 20"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(2)); // Third appointment is cancelled
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1)); // Only second appointment needs highlighting
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(2)); // Second row (after filtering)
            
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(result.Rows[1].Assistito, Is.EqualTo("Bianchi Luigi"));
        }

        [Test]
        public void Transform_WithDescrizionePuntoPartenza_DuplicatesText()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    DescrizionePuntoPartenza = "Casa",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "Via Verdi 10"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("CasaCasa"));
        }

        [Test]
        public void Transform_WithEmptyIndirizzoDestinazione_UsesCausaleDestinazione()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "",
                    CausaleDestinazione = "Ospedale"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Milano Ospedale"));
        }

        [Test]
        public void Transform_VerifiesEmptyColumns_AreAllEmpty()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    ComuneDestinazione = "Milano",
                    IndirizzoDestinazione = "Via Verdi 10"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            
            var row = result.Rows[0];
            Assert.That(row.EmptyColumn1, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn2, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn3, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn4, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn5, Is.EqualTo(string.Empty));
            Assert.That(row.Partenza, Is.EqualTo(string.Empty));
        }

        [Test]
        public void Transform_WithItalianCharacters_PreservesCharacters()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "09:00",
                    Attivita = "Trasporto",
                    DescrizioneStatoServizio = "Confermato",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "José",
                    ComuneDestinazione = "Città di Castello",
                    IndirizzoDestinazione = "Via dell'Università 5",
                    NoteERichieste = "Attenzione: è necessario aiuto"
                }
            };

            // Act
            var result = _dataTransformer.Transform(appointments);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            
            var row = result.Rows[0];
            Assert.That(row.Assistito, Is.EqualTo("Rossi José"));
            Assert.That(row.NomeAssistito, Is.EqualTo("José"));
            Assert.That(row.Indirizzo, Is.EqualTo("Città di Castello Via dell'Università 5"));
            Assert.That(row.NoteERichieste, Is.EqualTo("Attenzione: è necessario aiuto"));
        }
    }
}
