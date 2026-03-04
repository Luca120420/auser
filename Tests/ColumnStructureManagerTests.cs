using System;
using System.Collections.Generic;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for ColumnStructureManager.
    /// Tests specific examples and expected behaviors of the ColumnStructureManager.
    /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
    /// </summary>
    [TestFixture]
    public class ColumnStructureManagerTests
    {
        private ColumnStructureManager _columnStructureManager = null!;

        [SetUp]
        public void Setup()
        {
            _columnStructureManager = new ColumnStructureManager();
        }

        /// <summary>
        /// Test that GetColumnHeaders returns exactly 14 columns.
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnHeaders_ReturnsExactly15Columns()
        {
            // Act
            var headers = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers.Count, Is.EqualTo(14), "Column headers should contain exactly 14 columns");
        }

        /// <summary>
        /// Test that GetColumnHeaders returns the correct column names in the correct order.
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnHeaders_ReturnsCorrectColumnNames()
        {
            // Act
            var headers = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers[0], Is.EqualTo("Data"), "Column 1 should be 'Data'");
            Assert.That(headers[1], Is.EqualTo("Partenza"), "Column 2 should be 'Partenza'");
            Assert.That(headers[2], Is.EqualTo("Assistito"), "Column 3 should be 'Assistito'");
            Assert.That(headers[3], Is.EqualTo("Indirizzo"), "Column 4 should be 'Indirizzo'");
            Assert.That(headers[4], Is.EqualTo("Destinazione"), "Column 5 should be 'Destinazione'");
            Assert.That(headers[5], Is.EqualTo("Note"), "Column 6 should be 'Note'");
            Assert.That(headers[6], Is.EqualTo("Auto"), "Column 7 should be 'Auto'");
            Assert.That(headers[7], Is.EqualTo("Volontario"), "Column 8 should be 'Volontario'");
            Assert.That(headers[8], Is.EqualTo("Arrivo"), "Column 9 should be 'Arrivo'");
            Assert.That(headers[9], Is.EqualTo("Avv"), "Column 10 should be 'Avv'");
            Assert.That(headers[10], Is.EqualTo(""), "Column 11 should be empty");
            Assert.That(headers[11], Is.EqualTo("Indirizzo Gasnet"), "Column 12 should be 'Indirizzo Gasnet'");
            Assert.That(headers[12], Is.EqualTo("Note Gasnet"), "Column 13 should be 'Note Gasnet'");
            Assert.That(headers[13], Is.EqualTo(""), "Column 14 should be empty");
        }

        /// <summary>
        /// Test GetColumnIndex for all named columns.
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnIndex_ReturnsCorrectIndexForAllColumns()
        {
            // Assert - Test all named columns
            Assert.That(_columnStructureManager.GetColumnIndex("Data"), Is.EqualTo(0), "Data should be at index 0");
            Assert.That(_columnStructureManager.GetColumnIndex("Partenza"), Is.EqualTo(1), "Partenza should be at index 1");
            Assert.That(_columnStructureManager.GetColumnIndex("Assistito"), Is.EqualTo(2), "Assistito should be at index 2");
            Assert.That(_columnStructureManager.GetColumnIndex("Indirizzo"), Is.EqualTo(3), "Indirizzo should be at index 3");
            Assert.That(_columnStructureManager.GetColumnIndex("Destinazione"), Is.EqualTo(4), "Destinazione should be at index 4");
            Assert.That(_columnStructureManager.GetColumnIndex("Note"), Is.EqualTo(5), "Note should be at index 5");
            Assert.That(_columnStructureManager.GetColumnIndex("Auto"), Is.EqualTo(6), "Auto should be at index 6");
            Assert.That(_columnStructureManager.GetColumnIndex("Volontario"), Is.EqualTo(7), "Volontario should be at index 7");
            Assert.That(_columnStructureManager.GetColumnIndex("Arrivo"), Is.EqualTo(8), "Arrivo should be at index 8");
            Assert.That(_columnStructureManager.GetColumnIndex("Avv"), Is.EqualTo(9), "Avv should be at index 9");
            Assert.That(_columnStructureManager.GetColumnIndex("Indirizzo Gasnet"), Is.EqualTo(11), "Indirizzo Gasnet should be at index 11");
            Assert.That(_columnStructureManager.GetColumnIndex("Note Gasnet"), Is.EqualTo(12), "Note Gasnet should be at index 12");
        }

        /// <summary>
        /// Test GetColumnIndex returns -1 for non-existent columns.
        /// Validates: Requirements 3.1
        /// </summary>
        [Test]
        public void GetColumnIndex_ReturnsNegativeOneForNonExistentColumn()
        {
            // Act
            int index = _columnStructureManager.GetColumnIndex("Comune Partenza");

            // Assert
            Assert.That(index, Is.EqualTo(-1), "Non-existent column should return -1");
        }

        /// <summary>
        /// Test GetColumnIndex returns -1 for null or empty column names.
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnIndex_ReturnsNegativeOneForNullOrEmptyColumnName()
        {
            // Assert
            Assert.That(_columnStructureManager.GetColumnIndex(null!), Is.EqualTo(-1), "Null column name should return -1");
            Assert.That(_columnStructureManager.GetColumnIndex(""), Is.EqualTo(-1), "Empty column name should return -1");
            Assert.That(_columnStructureManager.GetColumnIndex("   "), Is.EqualTo(-1), "Whitespace column name should return -1");
        }

        /// <summary>
        /// Test GetNewColumnName for renamed column "Ora Inizio Servizio" to "Partenza".
        /// Validates: Requirements 1.1
        /// </summary>
        [Test]
        public void GetNewColumnName_RenamesOraInizioServizioToPartenza()
        {
            // Act
            string newName = _columnStructureManager.GetNewColumnName("Ora Inizio Servizio");

            // Assert
            Assert.That(newName, Is.EqualTo("Partenza"), "Ora Inizio Servizio should be renamed to Partenza");
        }

        /// <summary>
        /// Test GetNewColumnName for "Indirizzo Partenza" keeps the same name.
        /// Validates: Requirements 1.2
        /// </summary>
        [Test]
        public void GetNewColumnName_KeepsIndirizzoPartenzaName()
        {
            // Act
            string newName = _columnStructureManager.GetNewColumnName("Indirizzo Partenza");

            // Assert
            Assert.That(newName, Is.EqualTo("Indirizzo Partenza"), "Indirizzo Partenza should keep the same name");
        }

        /// <summary>
        /// Test GetNewColumnName returns original name for unmapped columns.
        /// Validates: Requirements 1.1, 1.2
        /// </summary>
        [Test]
        public void GetNewColumnName_ReturnsOriginalNameForUnmappedColumns()
        {
            // Act
            string newName1 = _columnStructureManager.GetNewColumnName("Data");
            string newName2 = _columnStructureManager.GetNewColumnName("Assistito");
            string newName3 = _columnStructureManager.GetNewColumnName("Note");

            // Assert
            Assert.That(newName1, Is.EqualTo("Data"), "Data should keep the same name");
            Assert.That(newName2, Is.EqualTo("Assistito"), "Assistito should keep the same name");
            Assert.That(newName3, Is.EqualTo("Note"), "Note should keep the same name");
        }

        /// <summary>
        /// Test GetNewColumnName handles null or empty input gracefully.
        /// Validates: Requirements 1.1, 1.2
        /// </summary>
        [Test]
        public void GetNewColumnName_HandlesNullOrEmptyInput()
        {
            // Assert
            Assert.That(_columnStructureManager.GetNewColumnName(null!), Is.Null, "Null input should return null");
            Assert.That(_columnStructureManager.GetNewColumnName(""), Is.EqualTo(""), "Empty input should return empty");
            Assert.That(_columnStructureManager.GetNewColumnName("   "), Is.EqualTo("   "), "Whitespace input should return whitespace");
        }

        /// <summary>
        /// Test that GetColumnHeaders returns a new list instance (not a reference to internal list).
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnHeaders_ReturnsNewListInstance()
        {
            // Act
            var headers1 = _columnStructureManager.GetColumnHeaders();
            var headers2 = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers1, Is.Not.SameAs(headers2), "GetColumnHeaders should return a new list instance each time");
            Assert.That(headers1, Is.EqualTo(headers2), "Both lists should have the same content");
        }

        /// <summary>
        /// Test that modifying the returned list doesn't affect the internal state.
        /// Validates: Requirements 1.1, 1.2, 2.1, 3.1, 3.2, 7.1, 8.1
        /// </summary>
        [Test]
        public void GetColumnHeaders_ModifyingReturnedListDoesNotAffectInternalState()
        {
            // Act
            var headers1 = _columnStructureManager.GetColumnHeaders();
            headers1[0] = "Modified";
            var headers2 = _columnStructureManager.GetColumnHeaders();

            // Assert
            Assert.That(headers2[0], Is.EqualTo("Data"), "Internal state should not be affected by modifying returned list");
        }
    }
}
