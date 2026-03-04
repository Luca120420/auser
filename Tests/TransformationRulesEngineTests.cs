using System;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for TransformationRulesEngine class.
    /// Tests all 11 transformation rules for converting CSV data to Excel format.
    /// Validates: Requirements 4.1, 4.2, 4.3, 4.4, 4.5, 4.6, 4.7, 4.8, 4.9
    /// </summary>
    [TestFixture]
    public class TransformationRulesEngineTests
    {
        private TransformationRulesEngine _engine = null!;

        [SetUp]
        public void Setup()
        {
            _engine = new TransformationRulesEngine();
        }

        /// <summary>
        /// Helper method to create a basic service appointment with required fields
        /// </summary>
        private ServiceAppointment CreateBasicAppointment()
        {
            return new ServiceAppointment
            {
                DataServizio = "26/01/2026",
                OraInizioServizio = "09:00",
                Attivita = "Accompagnamento",
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
            };
        }

        // Rule 1 Tests: Yellow Highlighting
        // Validates: Requirements 4.1

        [Test]
        public void Transform_WithAccompagConMacchinaAttrezzata_MarksRowForYellowHighlight()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.Attivita = "Accompag. con macchina attrezzata";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.YellowHighlightRows, Is.Not.Empty);
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(1));
        }

        [Test]
        public void Transform_WithoutAccompagConMacchinaAttrezzata_DoesNotMarkForYellowHighlight()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.Attivita = "Accompagnamento normale";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }

        [Test]
        public void Transform_WithPartialMatchAccompagConMacchinaAttrezzata_MarksRowForYellowHighlight()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.Attivita = "Servizio di Accompag. con macchina attrezzata speciale";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.YellowHighlightRows, Is.Not.Empty);
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1));
        }

        [Test]
        public void Transform_WithMultipleRowsAndOneHighlight_MarksCorrectRow()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.Attivita = "Accompagnamento normale";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.Attivita = "Accompag. con macchina attrezzata";
            
            var appointment3 = CreateBasicAppointment();
            appointment3.Attivita = "Trasporto";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(2)); // Second row
        }

        [Test]
        public void Transform_WithEmptyAttivita_DoesNotMarkForYellowHighlight()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.Attivita = null;
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }

        // Rule 3 Tests: Filter out ANNULLATO rows
        // Validates: Requirements 4.3

        [Test]
        public void Transform_WithAnnullatoStatus_FiltersOutRow()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizioneStatoServizio = "ANNULLATO";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
        }

        [Test]
        public void Transform_WithAnnullatoCaseInsensitive_FiltersOutRow()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizioneStatoServizio = "annullato";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
        }

        [Test]
        public void Transform_WithMixedStatusIncludingAnnullato_FiltersOnlyAnnullato()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.DescrizioneStatoServizio = "Confermato";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment3 = CreateBasicAppointment();
            appointment3.DescrizioneStatoServizio = "In attesa";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(2));
        }

        [Test]
        public void Transform_WithAllAnnullato_ReturnsEmptyResult()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.DescrizioneStatoServizio = "annullato";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
        }

        // Rule 4 Tests: Duplicate DESCRIZIONE PUNTO PARTENZA
        // Validates: Requirements 4.4

        [Test]
        public void Transform_WithDescrizionePuntoPartenza_DuplicatesText()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = "Casa";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("CasaCasa"));
        }

        [Test]
        public void Transform_WithEmptyDescrizionePuntoPartenza_DoesNotDuplicate()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = "";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo(""));
        }

        [Test]
        public void Transform_WithNullDescrizionePuntoPartenza_DoesNotDuplicate()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = null;
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Destinazione, Is.Null);
        }

        [Test]
        public void Transform_WithLongDescrizionePuntoPartenza_DuplicatesCorrectly()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = "Residenza Anziani Via Roma";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("Residenza Anziani Via RomaResidenza Anziani Via Roma"));
        }

        // Rule 5 Tests: Create ASSISTITO column
        // Validates: Requirements 4.5

        [Test]
        public void Transform_CreatesAssistitoColumn_WithCognomeAndNome()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "Rossi";
            appointment.NomeAssistito = "Mario";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Rossi Mario"));
        }

        [Test]
        public void Transform_AssistitoColumn_PreservesItalianCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "D'Àngelo";
            appointment.NomeAssistito = "José";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("D'Àngelo José"));
        }

        [Test]
        public void Transform_AssistitoColumn_HandlesLongNames()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "Della Rovere";
            appointment.NomeAssistito = "Maria Antonietta";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Della Rovere Maria Antonietta"));
        }

        // Rule 6-7 Tests: Create INDIRIZZO column with conditional logic
        // Validates: Requirements 4.6, 4.7

        [Test]
        public void Transform_WithIndirizzoDestinazione_ConcatenatesComuneAndIndirizzo()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Milano";
            appointment.IndirizzoDestinazione = "Via Verdi 10";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Milano Via Verdi 10"));
        }

        [Test]
        public void Transform_WithEmptyIndirizzoDestinazione_ConcatenatesComuneAndCausale()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Milano";
            appointment.IndirizzoDestinazione = "";
            appointment.CausaleDestinazione = "Ospedale";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Milano Ospedale"));
        }

        [Test]
        public void Transform_WithNullIndirizzoDestinazione_ConcatenatesComuneAndCausale()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Milano";
            appointment.IndirizzoDestinazione = null;
            appointment.CausaleDestinazione = "Clinica";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Milano Clinica"));
        }

        [Test]
        public void Transform_WithNullComuneDestinazione_HandlesGracefully()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = null;
            appointment.IndirizzoDestinazione = "Via Verdi 10";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo(" Via Verdi 10"));
        }

        [Test]
        public void Transform_WithBothIndirizzoAndCausaleEmpty_ConcatenatesOnlyComune()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Milano";
            appointment.IndirizzoDestinazione = "";
            appointment.CausaleDestinazione = "";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Milano "));
        }

        // Rule 8 Tests: Column structure with 5 empty columns
        // Validates: Requirements 4.8

        [Test]
        public void Transform_CreatesCorrectColumnStructure()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            var row = result.Rows[0];
            Assert.That(row.DataServizio, Is.Not.Null);
            Assert.That(row.OraInizioServizio, Is.Not.Null);
            Assert.That(row.Assistito, Is.Not.Null);
            Assert.That(row.CognomeAssistito, Is.Not.Null);
            Assert.That(row.NomeAssistito, Is.Not.Null);
            Assert.That(row.Indirizzo, Is.Not.Null);
            Assert.That(row.Destinazione, Is.Not.Null);
            Assert.That(row.EmptyColumn1, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn2, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn3, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn4, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn5, Is.EqualTo(string.Empty));
            Assert.That(row.OraInizioServizioCopy, Is.Not.Null);
            Assert.That(row.Partenza, Is.Not.Null);
            Assert.That(row.NoteERichieste, Is.Not.Null);
        }

        [Test]
        public void Transform_AllFiveEmptyColumnsAreEmpty()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            var row = result.Rows[0];
            Assert.That(row.EmptyColumn1, Is.Empty);
            Assert.That(row.EmptyColumn2, Is.Empty);
            Assert.That(row.EmptyColumn3, Is.Empty);
            Assert.That(row.EmptyColumn4, Is.Empty);
            Assert.That(row.EmptyColumn5, Is.Empty);
        }

        // Rule 9 Tests: Copy ORA INIZIO SERVIZIO
        // Validates: Requirements 4.8

        [Test]
        public void Transform_CopiesOraInizioServizio()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.OraInizioServizio = "14:30";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].OraInizioServizio, Is.EqualTo("14:30"));
            Assert.That(result.Rows[0].OraInizioServizioCopy, Is.EqualTo("14:30"));
        }

        [Test]
        public void Transform_OraInizioServizioCopy_MatchesOriginal()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.OraInizioServizio = "09:15";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].OraInizioServizioCopy, Is.EqualTo(result.Rows[0].OraInizioServizio));
        }

        // Rule 10 Tests: Leave Partenza field empty
        // Validates: Requirements 4.9

        [Test]
        public void Transform_PartenzaFieldIsAlwaysEmpty()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].Partenza, Is.Empty);
        }

        [Test]
        public void Transform_MultipleRows_AllPartenzaFieldsAreEmpty()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            var appointment2 = CreateBasicAppointment();
            var appointment3 = CreateBasicAppointment();
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.All(r => r.Partenza == string.Empty), Is.True);
        }

        // Rule 11 Tests: Preserve NOTE E RICHIESTE
        // Validates: Requirements 4.8

        [Test]
        public void Transform_PreservesNoteERichieste()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.NoteERichieste = "Urgente - Portare sedia a rotelle";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].NoteERichieste, Is.EqualTo("Urgente - Portare sedia a rotelle"));
        }

        [Test]
        public void Transform_PreservesNoteERichieste_WithItalianCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.NoteERichieste = "È più urgente - Qualità superiore";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].NoteERichieste, Is.EqualTo("È più urgente - Qualità superiore"));
        }

        [Test]
        public void Transform_PreservesEmptyNoteERichieste()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.NoteERichieste = null;
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].NoteERichieste, Is.Null);
        }

        // Integration Tests: Multiple rules together

        [Test]
        public void Transform_WithCompleteAppointment_AppliesAllRules()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "26/01/2026",
                OraInizioServizio = "09:00",
                Attivita = "Accompag. con macchina attrezzata",
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
            };
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(1));
            
            var row = result.Rows[0];
            Assert.That(row.DataServizio, Is.EqualTo("26/01/2026"));
            Assert.That(row.OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(row.Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(row.CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(row.NomeAssistito, Is.EqualTo("Mario"));
            Assert.That(row.Indirizzo, Is.EqualTo("Milano Via Verdi 10"));
            Assert.That(row.Destinazione, Is.EqualTo("CasaCasa"));
            Assert.That(row.OraInizioServizioCopy, Is.EqualTo("09:00"));
            Assert.That(row.Partenza, Is.Empty);
            Assert.That(row.NoteERichieste, Is.EqualTo("Nessuna nota"));
        }

        [Test]
        public void Transform_WithEmptyList_ReturnsEmptyResult()
        {
            // Arrange
            var appointments = new List<ServiceAppointment>();

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }

        [Test]
        public void Transform_WithNullList_ThrowsArgumentNullException()
        {
            // Arrange
            List<ServiceAppointment>? appointments = null;

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _engine.Transform(appointments!));
        }

        [Test]
        public void Transform_WithMultipleAppointments_MaintainsCorrectRowIndexing()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.Attivita = "Normale";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.Attivita = "Accompag. con macchina attrezzata";
            
            var appointment3 = CreateBasicAppointment();
            appointment3.DescrizioneStatoServizio = "ANNULLATO"; // This should be filtered
            
            var appointment4 = CreateBasicAppointment();
            appointment4.Attivita = "Accompag. con macchina attrezzata";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3, appointment4 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(3)); // One filtered out
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(2));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(2)); // Second row
            Assert.That(result.YellowHighlightRows[1], Is.EqualTo(3)); // Fourth row becomes third after filtering
        }

        [Test]
        public void Transform_PreservesDataServizio()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DataServizio = "15/03/2026";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].DataServizio, Is.EqualTo("15/03/2026"));
        }

        [Test]
        public void Transform_MaintainsSeparateCognomeAndNomeColumns()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "Verdi";
            appointment.NomeAssistito = "Giuseppe";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows[0].CognomeAssistito, Is.EqualTo("Verdi"));
            Assert.That(result.Rows[0].NomeAssistito, Is.EqualTo("Giuseppe"));
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Verdi Giuseppe"));
        }

        // Edge Case Tests
        // Task 6.3: Write unit tests for transformation edge cases
        // Validates: Requirements 4.4, 4.6, 4.7

        [Test]
        public void EdgeCase_EmptyDescrizionePuntoPartenza_ReturnsEmptyDestinazione()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = "";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo(""));
        }

        [Test]
        public void EdgeCase_NullDescrizionePuntoPartenza_ReturnsNullDestinazione()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = null;
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Destinazione, Is.Null);
        }

        [Test]
        public void EdgeCase_WhitespaceDescrizionePuntoPartenza_DuplicatesWhitespace()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.DescrizionePuntoPartenza = "   ";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("      ")); // Whitespace duplicated
        }

        [Test]
        public void EdgeCase_EmptyIndirizzoDestinazione_UsesCausaleDestinazione()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Roma";
            appointment.IndirizzoDestinazione = "";
            appointment.CausaleDestinazione = "Clinica San Raffaele";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Roma Clinica San Raffaele"));
        }

        [Test]
        public void EdgeCase_NullIndirizzoDestinazione_UsesCausaleDestinazione()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Napoli";
            appointment.IndirizzoDestinazione = null;
            appointment.CausaleDestinazione = "Ospedale Cardarelli";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Napoli Ospedale Cardarelli"));
        }

        [Test]
        public void EdgeCase_EmptyIndirizzoAndEmptyCausale_ConcatenatesOnlyComune()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Torino";
            appointment.IndirizzoDestinazione = "";
            appointment.CausaleDestinazione = "";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Torino "));
        }

        [Test]
        public void EdgeCase_NullIndirizzoAndNullCausale_ConcatenatesOnlyComune()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Firenze";
            appointment.IndirizzoDestinazione = null;
            appointment.CausaleDestinazione = null;
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Firenze "));
        }

        [Test]
        public void EdgeCase_AllAppointmentsCancelled_ReturnsEmptyResult()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.DescrizioneStatoServizio = "annullato";
            
            var appointment3 = CreateBasicAppointment();
            appointment3.DescrizioneStatoServizio = "Annullato";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }

        [Test]
        public void EdgeCase_AllAppointmentsCancelledWithMixedStatuses_ReturnsEmptyResult()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment2 = CreateBasicAppointment();
            appointment2.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment3 = CreateBasicAppointment();
            appointment3.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointment4 = CreateBasicAppointment();
            appointment4.DescrizioneStatoServizio = "ANNULLATO";
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2, appointment3, appointment4 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }

        [Test]
        public void EdgeCase_SpecialItalianCharactersInNames_PreservesCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "D'Àngelo";
            appointment.NomeAssistito = "José";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].CognomeAssistito, Is.EqualTo("D'Àngelo"));
            Assert.That(result.Rows[0].NomeAssistito, Is.EqualTo("José"));
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("D'Àngelo José"));
        }

        [Test]
        public void EdgeCase_SpecialItalianCharactersWithAccents_PreservesAllAccents()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "Pérez";
            appointment.NomeAssistito = "María";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].CognomeAssistito, Is.EqualTo("Pérez"));
            Assert.That(result.Rows[0].NomeAssistito, Is.EqualTo("María"));
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Pérez María"));
        }

        [Test]
        public void EdgeCase_ItalianCharactersInAddresses_PreservesCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.ComuneDestinazione = "Città di Castello";
            appointment.IndirizzoDestinazione = "Via dell'Università";
            appointment.DescrizionePuntoPartenza = "Piazza Libertà";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Città di Castello Via dell'Università"));
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("Piazza LibertàPiazza Libertà"));
        }

        [Test]
        public void EdgeCase_ComplexItalianCharactersInNotes_PreservesCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.NoteERichieste = "È necessario più tempo. L'assistito è fragile.";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].NoteERichieste, Is.EqualTo("È necessario più tempo. L'assistito è fragile."));
        }

        [Test]
        public void EdgeCase_AllItalianAccentedVowels_PreservesAllCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "Àlberto";
            appointment.NomeAssistito = "Èlena";
            appointment.NoteERichieste = "Ìsola, Òpera, Ùmbria - àèéìòù";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].CognomeAssistito, Is.EqualTo("Àlberto"));
            Assert.That(result.Rows[0].NomeAssistito, Is.EqualTo("Èlena"));
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("Àlberto Èlena"));
            Assert.That(result.Rows[0].NoteERichieste, Is.EqualTo("Ìsola, Òpera, Ùmbria - àèéìòù"));
        }

        [Test]
        public void EdgeCase_CombinedEdgeCases_EmptyFieldsAndItalianCharacters()
        {
            // Arrange
            var appointment = CreateBasicAppointment();
            appointment.CognomeAssistito = "D'Amico";
            appointment.NomeAssistito = "Françoise";
            appointment.DescrizionePuntoPartenza = ""; // Empty
            appointment.IndirizzoDestinazione = null; // Null
            appointment.CausaleDestinazione = "Ospedale Città";
            appointment.ComuneDestinazione = "Città di Castello";
            var appointments = new List<ServiceAppointment> { appointment };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.Rows[0].Assistito, Is.EqualTo("D'Amico Françoise"));
            Assert.That(result.Rows[0].Destinazione, Is.EqualTo("")); // Empty because DescrizionePuntoPartenza is empty
            Assert.That(result.Rows[0].Indirizzo, Is.EqualTo("Città di Castello Ospedale Città")); // Uses Causale because Indirizzo is null
        }

        [Test]
        public void EdgeCase_AllCancelledWithYellowHighlightCandidate_ReturnsEmptyWithNoHighlights()
        {
            // Arrange
            var appointment1 = CreateBasicAppointment();
            appointment1.Attivita = "Accompag. con macchina attrezzata"; // Would be highlighted
            appointment1.DescrizioneStatoServizio = "ANNULLATO"; // But is cancelled
            
            var appointment2 = CreateBasicAppointment();
            appointment2.Attivita = "Accompag. con macchina attrezzata"; // Would be highlighted
            appointment2.DescrizioneStatoServizio = "annullato"; // But is cancelled
            
            var appointments = new List<ServiceAppointment> { appointment1, appointment2 };

            // Act
            var result = _engine.Transform(appointments);

            // Assert
            Assert.That(result.Rows, Is.Empty);
            Assert.That(result.YellowHighlightRows, Is.Empty);
        }
    }
}
