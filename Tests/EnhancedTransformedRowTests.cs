using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the EnhancedTransformedRow model class.
    /// Validates: Requirements 1.1, 1.2, 2.1, 5.1, 6.1, 7.1, 7.2, 8.1, 8.2
    /// </summary>
    [TestFixture]
    public class EnhancedTransformedRowTests
    {
        [Test]
        public void EnhancedTransformedRow_CanBeInstantiated()
        {
            // Arrange & Act
            var row = new EnhancedTransformedRow();

            // Assert
            Assert.That(row, Is.Not.Null);
        }

        [Test]
        public void EnhancedTransformedRow_AllPropertiesCanBeSet()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Partenza = "09:00",
                Assistito = "Rossi Mario",
                Indirizzo = "Via Roma 10",
                Destinazione = "Ospedale",
                Note = "Portare documenti",
                Auto = "Fiat Panda",
                Volontario = "Bianchi Giuseppe",
                Arrivo = "09:30",
                Avv = "SI",
                Empty1 = "",
                IndirizzoGasnet = "Milano Via Verdi 10",
                NoteGasnet = "Note da CSV",
                Empty2 = ""
            };

            // Assert
            Assert.That(row.Data, Is.EqualTo("2024-01-26"));
            Assert.That(row.Partenza, Is.EqualTo("09:00"));
            Assert.That(row.Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(row.Indirizzo, Is.EqualTo("Via Roma 10"));
            Assert.That(row.Destinazione, Is.EqualTo("Ospedale"));
            Assert.That(row.Note, Is.EqualTo("Portare documenti"));
            Assert.That(row.Auto, Is.EqualTo("Fiat Panda"));
            Assert.That(row.Volontario, Is.EqualTo("Bianchi Giuseppe"));
            Assert.That(row.Arrivo, Is.EqualTo("09:30"));
            Assert.That(row.Avv, Is.EqualTo("SI"));
            Assert.That(row.Empty1, Is.EqualTo(""));
            Assert.That(row.IndirizzoGasnet, Is.EqualTo("Milano Via Verdi 10"));
            Assert.That(row.NoteGasnet, Is.EqualTo("Note da CSV"));
            Assert.That(row.Empty2, Is.EqualTo(""));
        }

        [Test]
        public void EnhancedTransformedRow_HasCorrectColumnStructure()
        {
            // Arrange
            var row = new EnhancedTransformedRow();
            var properties = typeof(EnhancedTransformedRow).GetProperties();

            // Act & Assert - Verify the 14-column structure
            var propertyNames = properties.Select(p => p.Name).ToList();
            
            Assert.That(propertyNames.Count, Is.EqualTo(14), "Should have exactly 14 columns");
            Assert.That(propertyNames[0], Is.EqualTo("Data"));
            Assert.That(propertyNames[1], Is.EqualTo("Partenza"));
            Assert.That(propertyNames[2], Is.EqualTo("Assistito"));
            Assert.That(propertyNames[3], Is.EqualTo("Indirizzo"));
            Assert.That(propertyNames[4], Is.EqualTo("Destinazione"));
            Assert.That(propertyNames[5], Is.EqualTo("Note"));
            Assert.That(propertyNames[6], Is.EqualTo("Auto"));
            Assert.That(propertyNames[7], Is.EqualTo("Volontario"));
            Assert.That(propertyNames[8], Is.EqualTo("Arrivo"));
            Assert.That(propertyNames[9], Is.EqualTo("Avv"));
            Assert.That(propertyNames[10], Is.EqualTo("Empty1"));
            Assert.That(propertyNames[11], Is.EqualTo("IndirizzoGasnet"));
            Assert.That(propertyNames[12], Is.EqualTo("NoteGasnet"));
            Assert.That(propertyNames[13], Is.EqualTo("Empty2"));
        }

        [Test]
        public void EnhancedTransformedRow_IndirizzoPositionedAfterAssistito()
        {
            // Arrange
            var properties = typeof(EnhancedTransformedRow).GetProperties();
            var propertyNames = properties.Select(p => p.Name).ToList();

            // Act
            int assistitoIndex = propertyNames.IndexOf("Assistito");
            int indirizzoIndex = propertyNames.IndexOf("Indirizzo");

            // Assert - Validates Requirement 2.1
            Assert.That(indirizzoIndex, Is.EqualTo(assistitoIndex + 1), 
                "Indirizzo should be positioned immediately after Assistito");
        }

        [Test]
        public void EnhancedTransformedRow_HasTwoEmptyColumns()
        {
            // Arrange
            var row = new EnhancedTransformedRow();

            // Act & Assert - Verify both empty columns exist and are initialized to empty string
            Assert.That(row.Empty1, Is.EqualTo(string.Empty));
            Assert.That(row.Empty2, Is.EqualTo(string.Empty));
        }

        [Test]
        public void EnhancedTransformedRow_PreservesItalianCharacters()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Partenza = "09:00",
                Assistito = "Àgostini Nicolò",
                Indirizzo = "Città di Castello Piazza dell'Università",
                Destinazione = "Università",
                Note = "Attenzione: è necessario l'accompagnatore",
                Volontario = "Pérez José",
                NoteGasnet = "Note con caratteri speciali: à è é ì ò ù"
            };

            // Assert - Verify Italian characters are preserved
            Assert.That(row.Assistito, Is.EqualTo("Àgostini Nicolò"));
            Assert.That(row.Indirizzo, Is.EqualTo("Città di Castello Piazza dell'Università"));
            Assert.That(row.Destinazione, Is.EqualTo("Università"));
            Assert.That(row.Note, Is.EqualTo("Attenzione: è necessario l'accompagnatore"));
            Assert.That(row.Volontario, Is.EqualTo("Pérez José"));
            Assert.That(row.NoteGasnet, Is.EqualTo("Note con caratteri speciali: à è é ì ò ù"));
        }

        [Test]
        public void EnhancedTransformedRow_RequiredFields_Data()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                // Data is empty (default value)
                Assistito = "Rossi Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("Data")), Is.True);
        }

        [Test]
        public void EnhancedTransformedRow_RequiredFields_Assistito()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26"
                // Assistito is empty (default value)
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("Assistito")), Is.True);
        }

        [Test]
        public void EnhancedTransformedRow_AllRequiredFieldsPresent_IsValid()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void EnhancedTransformedRow_OptionalFieldsCanBeEmpty()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario",
                // Optional fields left as empty
                Partenza = "",
                Indirizzo = "",
                Destinazione = "",
                Note = "",
                Auto = "",
                Volontario = "",
                Arrivo = "",
                Avv = "",
                IndirizzoGasnet = "",
                NoteGasnet = ""
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
        }

        [Test]
        public void EnhancedTransformedRow_EmptyColumnsRemainEmpty()
        {
            // Arrange
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Partenza = "09:00",
                Assistito = "Rossi Mario",
                Indirizzo = "Via Roma 10",
                Destinazione = "Ospedale",
                Note = "Test",
                Auto = "Fiat",
                Volontario = "Bianchi",
                Arrivo = "09:30",
                Avv = "SI",
                IndirizzoGasnet = "Milano",
                NoteGasnet = "Note"
            };

            // Assert - Verify empty columns remain empty
            Assert.That(row.Empty1, Is.Empty);
            Assert.That(row.Empty2, Is.Empty);
        }

        [Test]
        public void EnhancedTransformedRow_PartenzaColumn_RenamedFromOraInizioServizio()
        {
            // Arrange
            var properties = typeof(EnhancedTransformedRow).GetProperties();
            var propertyNames = properties.Select(p => p.Name).ToList();

            // Assert - Validates Requirement 1.1 (column renamed)
            Assert.That(propertyNames.Contains("Partenza"), Is.True, 
                "Should have Partenza column (renamed from Ora Inizio Servizio)");
            Assert.That(propertyNames.Contains("OraInizioServizio"), Is.False, 
                "Should not have OraInizioServizio column (renamed to Partenza)");
        }

        [Test]
        public void EnhancedTransformedRow_NoteGasnetColumn_ExistsAfterIndirizzoGasnet()
        {
            // Arrange
            var properties = typeof(EnhancedTransformedRow).GetProperties();
            var propertyNames = properties.Select(p => p.Name).ToList();

            // Act
            int indirizzoGasnetIndex = propertyNames.IndexOf("IndirizzoGasnet");
            int noteGasnetIndex = propertyNames.IndexOf("NoteGasnet");

            // Assert - Validates Requirement 7.1 (Note Gasnet positioned after Indirizzo Gasnet)
            Assert.That(noteGasnetIndex, Is.EqualTo(indirizzoGasnetIndex + 1), 
                "Note Gasnet should be positioned immediately after Indirizzo Gasnet");
        }

        [Test]
        public void EnhancedTransformedRow_AvvColumn_ExistsAfterArrivo()
        {
            // Arrange
            var properties = typeof(EnhancedTransformedRow).GetProperties();
            var propertyNames = properties.Select(p => p.Name).ToList();

            // Act
            int arrivoIndex = propertyNames.IndexOf("Arrivo");
            int avvIndex = propertyNames.IndexOf("Avv");

            // Assert - Validates Requirement 8.1 (Avv positioned after Arrivo)
            Assert.That(avvIndex, Is.EqualTo(arrivoIndex + 1), 
                "Avv should be positioned immediately after Arrivo");
        }

        [Test]
        public void EnhancedTransformedRow_NoteGasnetCanStoreCSVNotes()
        {
            // Arrange
            var csvNote = "Nota importante dal CSV con caratteri speciali: à è é";
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario",
                NoteGasnet = csvNote
            };

            // Assert - Validates Requirement 7.2 (Note Gasnet stores CSV note data)
            Assert.That(row.NoteGasnet, Is.EqualTo(csvNote));
        }

        [Test]
        public void EnhancedTransformedRow_IndirizzoCanStoreLookupResult()
        {
            // Arrange
            var lookupAddress = "Via Roma 10, Milano";
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario",
                Indirizzo = lookupAddress
            };

            // Assert - Validates Requirement 5.1 (Indirizzo stores assistiti lookup result)
            Assert.That(row.Indirizzo, Is.EqualTo(lookupAddress));
        }

        [Test]
        public void EnhancedTransformedRow_NoteCanStoreLookupResult()
        {
            // Arrange
            var lookupNote = "Note dal foglio assistiti";
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario",
                Note = lookupNote
            };

            // Assert - Validates Requirement 6.1 (Note stores assistiti lookup result)
            Assert.That(row.Note, Is.EqualTo(lookupNote));
        }

        [Test]
        public void EnhancedTransformedRow_AvvCanStoreFissiLookupResult()
        {
            // Arrange
            var lookupAvv = "SI";
            var row = new EnhancedTransformedRow
            {
                Data = "2024-01-26",
                Assistito = "Rossi Mario",
                Avv = lookupAvv
            };

            // Assert - Validates Requirement 8.2 (Avv stores fissi lookup result)
            Assert.That(row.Avv, Is.EqualTo(lookupAvv));
        }
    }
}
