using System;
using System.Collections.Generic;
using NUnit.Framework;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for LookupService edge cases and error conditions.
    /// Tests specific examples and edge cases for the LookupService.
    /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 8.2, 8.3
    /// </summary>
    [TestFixture]
    public class LookupServiceTests
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
        /// Helper method to create a test assistiti sheet with data
        /// </summary>
        private Sheet CreateAssistitiSheet(Dictionary<string, (string Indirizzo, string Note)> data)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");

            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";

            // Write data rows
            int row = 2;
            foreach (var kvp in data)
            {
                worksheet.Cells[row, 1].Value = kvp.Key;
                worksheet.Cells[row, 2].Value = kvp.Value.Indirizzo;
                worksheet.Cells[row, 3].Value = kvp.Value.Note;
                row++;
            }

            return new Sheet(worksheet);
        }

        /// <summary>
        /// Helper method to create a test fissi sheet with data
        /// </summary>
        private Sheet CreateFissiSheet(Dictionary<string, string> data)
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("fissi");

            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Avv";

            // Write data rows
            int row = 2;
            foreach (var kvp in data)
            {
                worksheet.Cells[row, 1].Value = kvp.Key;
                worksheet.Cells[row, 2].Value = kvp.Value;
                row++;
            }

            return new Sheet(worksheet);
        }

        /// <summary>
        /// Helper method to create an empty assistiti sheet (headers only)
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
        /// Helper method to create an empty fissi sheet (headers only)
        /// </summary>
        private Sheet CreateEmptyFissiSheet()
        {
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("fissi");
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Avv";
            return new Sheet(worksheet);
        }

        #region Empty Reference Sheet Tests

        /// <summary>
        /// Test that lookups on empty assistiti sheet return empty string
        /// </summary>
        [Test]
        public void LookupInAssistiti_EmptySheet_ReturnsEmptyString()
        {
            // Arrange
            var emptyAssistitiSheet = CreateEmptyAssistitiSheet();
            var emptyFissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(emptyAssistitiSheet, emptyFissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Mario Rossi", "Note");

            // Assert
            Assert.That(indirizzo, Is.EqualTo(""));
            Assert.That(note, Is.EqualTo(""));
        }

        /// <summary>
        /// Test that lookups on empty fissi sheet return empty string
        /// </summary>
        [Test]
        public void LookupInFissi_EmptySheet_ReturnsEmptyString()
        {
            // Arrange
            var emptyAssistitiSheet = CreateEmptyAssistitiSheet();
            var emptyFissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(emptyAssistitiSheet, emptyFissiSheet);

            // Act
            var avv = _lookupService.LookupInFissi("Mario Rossi", "Avv");

            // Assert
            Assert.That(avv, Is.EqualTo(""));
        }

        #endregion

        #region Special Characters Tests

        /// <summary>
        /// Test lookup with assistito name containing à character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÀ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Nicolà Bianchi", ("Via Roma 10", "Cliente abituale") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Nicolà Bianchi", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Nicolà Bianchi", "Note");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Via Roma 10"));
            Assert.That(note, Is.EqualTo("Cliente abituale"));
        }

        /// <summary>
        /// Test lookup with assistito name containing è character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÈ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Josè Verdi", ("Piazza Garibaldi 5", "Necessita assistenza") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Josè Verdi", "Indirizzo");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Piazza Garibaldi 5"));
        }

        /// <summary>
        /// Test lookup with assistito name containing é character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÉ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "André Neri", ("Corso Italia 20", "VIP") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("André Neri", "Indirizzo");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Corso Italia 20"));
        }

        /// <summary>
        /// Test lookup with assistito name containing ì character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÌ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Marìa Gialli", ("Via Dante 15", "Urgente") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Marìa Gialli", "Indirizzo");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Via Dante 15"));
        }

        /// <summary>
        /// Test lookup with assistito name containing ò character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÒ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Pòlo Azzurri", ("Viale Europa 30", "Standard") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Pòlo Azzurri", "Indirizzo");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Viale Europa 30"));
        }

        /// <summary>
        /// Test lookup with assistito name containing ù character
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithÙ_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Gesù Marroni", ("Via Torino 8", "Priorità alta") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Gesù Marroni", "Indirizzo");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Via Torino 8"));
        }

        /// <summary>
        /// Test lookup with assistito name containing multiple special characters
        /// </summary>
        [Test]
        public void LookupInAssistiti_NameWithMultipleSpecialChars_ReturnsCorrectValue()
        {
            // Arrange
            var data = new Dictionary<string, (string, string)>
            {
                { "Nicolò D'Àngelo", ("Via Mazzini 12", "Cliente speciale") }
            };
            var assistitiSheet = CreateAssistitiSheet(data);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Nicolò D'Àngelo", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Nicolò D'Àngelo", "Note");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("Via Mazzini 12"));
            Assert.That(note, Is.EqualTo("Cliente speciale"));
        }

        #endregion

        #region Multiple Matches Tests

        /// <summary>
        /// Test that when multiple matches exist, the first match is used
        /// </summary>
        [Test]
        public void LookupInAssistiti_MultipleMatches_UsesFirstMatch()
        {
            // Arrange - Create sheet with duplicate names
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            
            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";
            
            // Write duplicate entries (same name, different data)
            worksheet.Cells[2, 1].Value = "Mario Rossi";
            worksheet.Cells[2, 2].Value = "Via Roma 10";
            worksheet.Cells[2, 3].Value = "Prima occorrenza";
            
            worksheet.Cells[3, 1].Value = "Mario Rossi";
            worksheet.Cells[3, 2].Value = "Via Milano 20";
            worksheet.Cells[3, 3].Value = "Seconda occorrenza";
            
            worksheet.Cells[4, 1].Value = "Mario Rossi";
            worksheet.Cells[4, 2].Value = "Via Napoli 30";
            worksheet.Cells[4, 3].Value = "Terza occorrenza";

            var assistitiSheet = new Sheet(worksheet);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Mario Rossi", "Note");

            // Assert - Should return first match
            Assert.That(indirizzo, Is.EqualTo("Via Roma 10"));
            Assert.That(note, Is.EqualTo("Prima occorrenza"));
        }

        /// <summary>
        /// Test that when multiple matches exist in fissi sheet, the first match is used
        /// </summary>
        [Test]
        public void LookupInFissi_MultipleMatches_UsesFirstMatch()
        {
            // Arrange - Create sheet with duplicate names
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("fissi");
            
            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Avv";
            
            // Write duplicate entries (same name, different data)
            worksheet.Cells[2, 1].Value = "Luigi Verdi";
            worksheet.Cells[2, 2].Value = "Avviso 1";
            
            worksheet.Cells[3, 1].Value = "Luigi Verdi";
            worksheet.Cells[3, 2].Value = "Avviso 2";

            var fissiSheet = new Sheet(worksheet);
            var assistitiSheet = CreateEmptyAssistitiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var avv = _lookupService.LookupInFissi("Luigi Verdi", "Avv");

            // Assert - Should return first match
            Assert.That(avv, Is.EqualTo("Avviso 1"));
        }

        #endregion

        #region Malformed Reference Sheet Tests

        /// <summary>
        /// Test lookup on sheet with missing columns (no Indirizzo column)
        /// </summary>
        [Test]
        public void LookupInAssistiti_MissingColumn_ReturnsEmptyString()
        {
            // Arrange - Create sheet with only Nome and Note columns (missing Indirizzo)
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            
            // Write header row without Indirizzo column
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Note";
            
            // Write data row
            worksheet.Cells[2, 1].Value = "Mario Rossi";
            worksheet.Cells[2, 2].Value = "Cliente abituale";

            var assistitiSheet = new Sheet(worksheet);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Mario Rossi", "Note");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("")); // Missing column returns empty
            Assert.That(note, Is.EqualTo("Cliente abituale")); // Existing column works
        }

        /// <summary>
        /// Test lookup on sheet with no header row
        /// </summary>
        [Test]
        public void LookupInAssistiti_NoHeaderRow_HandlesGracefully()
        {
            // Arrange - Create sheet with data but no proper headers
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            
            // Write data directly without headers (or with data in first row)
            worksheet.Cells[1, 1].Value = "Mario Rossi";
            worksheet.Cells[1, 2].Value = "Via Roma 10";
            worksheet.Cells[1, 3].Value = "Note";

            var assistitiSheet = new Sheet(worksheet);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act - Try to lookup (first row is treated as header)
            var indirizzo = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");

            // Assert - Should return empty since first row is treated as header
            Assert.That(indirizzo, Is.EqualTo(""));
        }

        /// <summary>
        /// Test lookup on sheet with empty cells in data rows
        /// </summary>
        [Test]
        public void LookupInAssistiti_EmptyCellsInDataRow_ReturnsEmptyString()
        {
            // Arrange - Create sheet with empty cells
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            
            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";
            
            // Write data row with empty Indirizzo
            worksheet.Cells[2, 1].Value = "Mario Rossi";
            // worksheet.Cells[2, 2] is intentionally left empty
            worksheet.Cells[2, 3].Value = "Cliente abituale";

            var assistitiSheet = new Sheet(worksheet);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var indirizzo = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");
            var note = _lookupService.LookupInAssistiti("Mario Rossi", "Note");

            // Assert
            Assert.That(indirizzo, Is.EqualTo("")); // Empty cell returns empty string
            Assert.That(note, Is.EqualTo("Cliente abituale"));
        }

        /// <summary>
        /// Test lookup on sheet with rows that have empty lookup keys (should be skipped)
        /// </summary>
        [Test]
        public void LookupInAssistiti_EmptyLookupKey_SkipsRow()
        {
            // Arrange - Create sheet with empty lookup key in one row
            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("assistiti");
            
            // Write header row
            worksheet.Cells[1, 1].Value = "Nome";
            worksheet.Cells[1, 2].Value = "Indirizzo";
            worksheet.Cells[1, 3].Value = "Note";
            
            // Write row with empty Nome (lookup key)
            worksheet.Cells[2, 1].Value = ""; // Empty lookup key
            worksheet.Cells[2, 2].Value = "Via Roma 10";
            worksheet.Cells[2, 3].Value = "Should be skipped";
            
            // Write valid row
            worksheet.Cells[3, 1].Value = "Mario Rossi";
            worksheet.Cells[3, 2].Value = "Via Milano 20";
            worksheet.Cells[3, 3].Value = "Valid entry";

            var assistitiSheet = new Sheet(worksheet);
            var fissiSheet = CreateEmptyFissiSheet();
            _lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);

            // Act
            var emptyKeyResult = _lookupService.LookupInAssistiti("", "Indirizzo");
            var validResult = _lookupService.LookupInAssistiti("Mario Rossi", "Indirizzo");

            // Assert
            Assert.That(emptyKeyResult, Is.EqualTo("")); // Empty key returns empty
            Assert.That(validResult, Is.EqualTo("Via Milano 20")); // Valid key works
        }

        #endregion
    }
}
