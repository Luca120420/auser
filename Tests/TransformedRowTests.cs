using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the TransformedRow model class.
    /// Validates: Requirements 4.8
    /// </summary>
    [TestFixture]
    public class TransformedRowTests
    {
        [Test]
        public void TransformedRow_CanBeInstantiated()
        {
            // Arrange & Act
            var row = new TransformedRow();

            // Assert
            Assert.That(row, Is.Not.Null);
        }

        [Test]
        public void TransformedRow_AllPropertiesCanBeSet()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                Indirizzo = "Milano Via Verdi 10",
                Destinazione = "Ospedale",
                EmptyColumn1 = "",
                EmptyColumn2 = "",
                EmptyColumn3 = "",
                EmptyColumn4 = "",
                EmptyColumn5 = "",
                OraInizioServizioCopy = "09:00",
                Partenza = "",
                NoteERichieste = "Portare documenti"
            };

            // Assert
            Assert.That(row.DataServizio, Is.EqualTo("2024-01-26"));
            Assert.That(row.OraInizioServizio, Is.EqualTo("09:00"));
            Assert.That(row.Assistito, Is.EqualTo("Rossi Mario"));
            Assert.That(row.CognomeAssistito, Is.EqualTo("Rossi"));
            Assert.That(row.NomeAssistito, Is.EqualTo("Mario"));
            Assert.That(row.Indirizzo, Is.EqualTo("Milano Via Verdi 10"));
            Assert.That(row.Destinazione, Is.EqualTo("Ospedale"));
            Assert.That(row.EmptyColumn1, Is.EqualTo(""));
            Assert.That(row.EmptyColumn2, Is.EqualTo(""));
            Assert.That(row.EmptyColumn3, Is.EqualTo(""));
            Assert.That(row.EmptyColumn4, Is.EqualTo(""));
            Assert.That(row.EmptyColumn5, Is.EqualTo(""));
            Assert.That(row.OraInizioServizioCopy, Is.EqualTo("09:00"));
            Assert.That(row.Partenza, Is.EqualTo(""));
            Assert.That(row.NoteERichieste, Is.EqualTo("Portare documenti"));
        }

        [Test]
        public void TransformedRow_HasCorrectColumnStructure()
        {
            // Arrange
            var row = new TransformedRow();
            var properties = typeof(TransformedRow).GetProperties();

            // Act & Assert - Verify the column order matches requirements 4.8
            var propertyNames = properties.Select(p => p.Name).ToList();
            
            Assert.That(propertyNames[0], Is.EqualTo("DataServizio"));
            Assert.That(propertyNames[1], Is.EqualTo("OraInizioServizio"));
            Assert.That(propertyNames[2], Is.EqualTo("Assistito"));
            Assert.That(propertyNames[3], Is.EqualTo("CognomeAssistito"));
            Assert.That(propertyNames[4], Is.EqualTo("NomeAssistito"));
            Assert.That(propertyNames[5], Is.EqualTo("Indirizzo"));
            Assert.That(propertyNames[6], Is.EqualTo("Destinazione"));
            Assert.That(propertyNames[7], Is.EqualTo("EmptyColumn1"));
            Assert.That(propertyNames[8], Is.EqualTo("EmptyColumn2"));
            Assert.That(propertyNames[9], Is.EqualTo("EmptyColumn3"));
            Assert.That(propertyNames[10], Is.EqualTo("EmptyColumn4"));
            Assert.That(propertyNames[11], Is.EqualTo("EmptyColumn5"));
            Assert.That(propertyNames[12], Is.EqualTo("OraInizioServizioCopy"));
            Assert.That(propertyNames[13], Is.EqualTo("Partenza"));
            Assert.That(propertyNames[14], Is.EqualTo("NoteERichieste"));
        }

        [Test]
        public void TransformedRow_HasFiveEmptyColumns()
        {
            // Arrange
            var row = new TransformedRow();

            // Act & Assert - Verify all 5 empty columns exist and are initialized to empty string
            Assert.That(row.EmptyColumn1, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn2, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn3, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn4, Is.EqualTo(string.Empty));
            Assert.That(row.EmptyColumn5, Is.EqualTo(string.Empty));
        }

        [Test]
        public void TransformedRow_PartenzaFieldIsEmpty()
        {
            // Arrange & Act
            var row = new TransformedRow();

            // Assert - Verify Partenza field is initialized to empty string (Requirement 4.9)
            Assert.That(row.Partenza, Is.EqualTo(string.Empty));
        }

        [Test]
        public void TransformedRow_PreservesItalianCharacters()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Àgostini Nicolò",
                CognomeAssistito = "Àgostini",
                NomeAssistito = "Nicolò",
                Indirizzo = "Città di Castello Piazza dell'Università",
                Destinazione = "Università",
                OraInizioServizioCopy = "09:00",
                NoteERichieste = "Attenzione: è necessario l'accompagnatore"
            };

            // Assert - Verify Italian characters are preserved
            Assert.That(row.Assistito, Is.EqualTo("Àgostini Nicolò"));
            Assert.That(row.CognomeAssistito, Is.EqualTo("Àgostini"));
            Assert.That(row.NomeAssistito, Is.EqualTo("Nicolò"));
            Assert.That(row.Indirizzo, Is.EqualTo("Città di Castello Piazza dell'Università"));
            Assert.That(row.Destinazione, Is.EqualTo("Università"));
            Assert.That(row.NoteERichieste, Is.EqualTo("Attenzione: è necessario l'accompagnatore"));
        }

        [Test]
        public void TransformedRow_RequiredFields_DataServizio()
        {
            // Arrange
            var row = new TransformedRow
            {
                // DataServizio is empty (default value)
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("DataServizio")), Is.True);
        }

        [Test]
        public void TransformedRow_RequiredFields_OraInizioServizio()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                // OraInizioServizio is empty (default value)
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("OraInizioServizio")), Is.True);
        }

        [Test]
        public void TransformedRow_RequiredFields_Assistito()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                // Assistito is empty (default value)
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
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
        public void TransformedRow_RequiredFields_CognomeAssistito()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                // CognomeAssistito is empty (default value)
                NomeAssistito = "Mario"
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("CognomeAssistito")), Is.True);
        }

        [Test]
        public void TransformedRow_RequiredFields_NomeAssistito()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi"
                // NomeAssistito is empty (default value)
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("NomeAssistito")), Is.True);
        }

        [Test]
        public void TransformedRow_AllRequiredFieldsPresent_IsValid()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
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
        public void TransformedRow_OptionalFieldsCanBeNull()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                // Optional fields left as null or empty
                Destinazione = null,
                NoteERichieste = null
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(row);
            var isValid = Validator.TryValidateObject(row, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
        }

        [Test]
        public void TransformedRow_OraInizioServizioCopy_MatchesOriginal()
        {
            // Arrange
            var originalTime = "14:30";
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = originalTime,
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                OraInizioServizioCopy = originalTime
            };

            // Assert - Verify the copy matches the original
            Assert.That(row.OraInizioServizioCopy, Is.EqualTo(row.OraInizioServizio));
        }

        [Test]
        public void TransformedRow_EmptyColumnsRemainEmpty()
        {
            // Arrange
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario",
                Indirizzo = "Milano Via Verdi 10",
                Destinazione = "Ospedale",
                OraInizioServizioCopy = "09:00",
                NoteERichieste = "Test"
            };

            // Assert - Verify all empty columns remain empty (Requirement 4.8)
            Assert.That(row.EmptyColumn1, Is.Empty);
            Assert.That(row.EmptyColumn2, Is.Empty);
            Assert.That(row.EmptyColumn3, Is.Empty);
            Assert.That(row.EmptyColumn4, Is.Empty);
            Assert.That(row.EmptyColumn5, Is.Empty);
        }
    }
}
