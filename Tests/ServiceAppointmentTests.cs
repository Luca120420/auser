using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the ServiceAppointment model class.
    /// Validates: Requirements 2.2
    /// </summary>
    [TestFixture]
    public class ServiceAppointmentTests
    {
        [Test]
        public void ServiceAppointment_CanBeInstantiated()
        {
            // Arrange & Act
            var appointment = new ServiceAppointment();

            // Assert
            Assert.That(appointment, Is.Not.Null);
        }

        [Test]
        public void ServiceAppointment_AllPropertiesCanBeSet()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Attivita = "Accompag. con macchina attrezzata",
                DescrizioneStatoServizio = "CONFERMATO",
                IndirizzoPartenza = "Via Roma 1",
                ComunePartenza = "Milano",
                DescrizionePuntoPartenza = "Ingresso principale",
                IndirizzoDestinazione = "Via Verdi 10",
                ComuneDestinazione = "Milano",
                CausaleDestinazione = "Visita medica",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                NoteERichieste = "Portare documenti"
            };

            // Assert
            Assert.That(appointment.DataServizio, Is.EqualTo("2024-01-26"));
            Assert.That(appointment.OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(appointment.Attivita, Is.EqualTo("Accompag. con macchina attrezzata"));
            Assert.That(appointment.DescrizioneStatoServizio, Is.EqualTo("CONFERMATO"));
            Assert.That(appointment.IndirizzoPartenza, Is.EqualTo("Via Roma 1"));
            Assert.That(appointment.ComunePartenza, Is.EqualTo("Milano"));
            Assert.That(appointment.DescrizionePuntoPartenza, Is.EqualTo("Ingresso principale"));
            Assert.That(appointment.IndirizzoDestinazione, Is.EqualTo("Via Verdi 10"));
            Assert.That(appointment.ComuneDestinazione, Is.EqualTo("Milano"));
            Assert.That(appointment.CausaleDestinazione, Is.EqualTo("Visita medica"));
            Assert.That(appointment.CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(appointment.NomeAssistito, Is.EqualTo("Mario"));
            Assert.That(appointment.NoteERichieste, Is.EqualTo("Portare documenti"));
        }

        [Test]
        public void ServiceAppointment_PreservesItalianCharacters()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                CognomeAssistito = "Àgostini",
                NomeAssistito = "Nicolò",
                IndirizzoDestinazione = "Piazza dell'Università",
                ComuneDestinazione = "Città di Castello",
                NoteERichieste = "Attenzione: è necessario l'accompagnatore"
            };

            // Assert - Verify Italian characters are preserved
            Assert.That(appointment.CognomeAssistito, Is.EqualTo("Àgostini"));
            Assert.That(appointment.NomeAssistito, Is.EqualTo("Nicolò"));
            Assert.That(appointment.IndirizzoDestinazione, Is.EqualTo("Piazza dell'Università"));
            Assert.That(appointment.ComuneDestinazione, Is.EqualTo("Città di Castello"));
            Assert.That(appointment.NoteERichieste, Is.EqualTo("Attenzione: è necessario l'accompagnatore"));
        }

        [Test]
        public void ServiceAppointment_RequiredFields_DataServizio()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                // DataServizio is empty (default value)
                OraInizioServizio = "09:00",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("DataServizio")), Is.True);
        }

        [Test]
        public void ServiceAppointment_RequiredFields_OraInizioServizio()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                // OraInizioServizio is empty (default value)
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("OraInizioServizio")), Is.True);
        }

        [Test]
        public void ServiceAppointment_RequiredFields_CognomeAssistito()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                // CognomeAssistito is empty (default value)
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("CognomeAssistito")), Is.True);
        }

        [Test]
        public void ServiceAppointment_RequiredFields_NomeAssistito()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                CognomeAssistito = "Rossi"
                // NomeAssistito is empty (default value)
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("NomeAssistito")), Is.True);
        }

        [Test]
        public void ServiceAppointment_AllRequiredFieldsPresent_IsValid()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void ServiceAppointment_OptionalFieldsCanBeNull()
        {
            // Arrange
            var appointment = new ServiceAppointment
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                // Optional fields left as null
                Attivita = null,
                DescrizioneStatoServizio = null,
                IndirizzoPartenza = null,
                ComunePartenza = null,
                DescrizionePuntoPartenza = null,
                IndirizzoDestinazione = null,
                ComuneDestinazione = null,
                CausaleDestinazione = null,
                NoteERichieste = null
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(appointment);
            var isValid = Validator.TryValidateObject(appointment, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
        }
    }
}
