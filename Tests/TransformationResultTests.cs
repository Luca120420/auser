using System;
using System.ComponentModel.DataAnnotations;
using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the TransformationResult model class.
    /// Validates: Requirements 4.1
    /// </summary>
    [TestFixture]
    public class TransformationResultTests
    {
        [Test]
        public void TransformationResult_CanBeInstantiated()
        {
            // Arrange & Act
            var result = new TransformationResult();

            // Assert
            Assert.That(result, Is.Not.Null);
        }

        [Test]
        public void TransformationResult_RowsProperty_IsInitializedAsEmptyList()
        {
            // Arrange & Act
            var result = new TransformationResult();

            // Assert
            Assert.That(result.Rows, Is.Not.Null);
            Assert.That(result.Rows, Is.Empty);
            Assert.That(result.Rows, Is.InstanceOf<List<TransformedRow>>());
        }

        [Test]
        public void TransformationResult_YellowHighlightRowsProperty_IsInitializedAsEmptyList()
        {
            // Arrange & Act
            var result = new TransformationResult();

            // Assert
            Assert.That(result.YellowHighlightRows, Is.Not.Null);
            Assert.That(result.YellowHighlightRows, Is.Empty);
            Assert.That(result.YellowHighlightRows, Is.InstanceOf<List<int>>());
        }

        [Test]
        public void TransformationResult_CanAddTransformedRows()
        {
            // Arrange
            var result = new TransformationResult();
            var row1 = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };
            var row2 = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "10:00",
                Assistito = "Bianchi Anna",
                CognomeAssistito = "Bianchi",
                NomeAssistito = "Anna"
            };

            // Act
            result.Rows.Add(row1);
            result.Rows.Add(row2);

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(2));
            Assert.That(result.Rows[0], Is.EqualTo(row1));
            Assert.That(result.Rows[1], Is.EqualTo(row2));
        }

        [Test]
        public void TransformationResult_CanAddYellowHighlightRowIndices()
        {
            // Arrange
            var result = new TransformationResult();

            // Act
            result.YellowHighlightRows.Add(1);
            result.YellowHighlightRows.Add(3);
            result.YellowHighlightRows.Add(5);

            // Assert
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(3));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[1], Is.EqualTo(3));
            Assert.That(result.YellowHighlightRows[2], Is.EqualTo(5));
        }

        [Test]
        public void TransformationResult_YellowHighlightRows_CanBeEmpty()
        {
            // Arrange
            var result = new TransformationResult();
            var row = new TransformedRow
            {
                DataServizio = "2024-01-26",
                OraInizioServizio = "09:00",
                Assistito = "Rossi Mario",
                CognomeAssistito = "Rossi",
                NomeAssistito = "Mario"
            };

            // Act
            result.Rows.Add(row);
            // Don't add any yellow highlight rows

            // Assert
            Assert.That(result.Rows.Count, Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(0));
        }

        [Test]
        public void TransformationResult_YellowHighlightRows_UsesOneBasedIndexing()
        {
            // Arrange
            var result = new TransformationResult();
            
            // Act - Add row indices starting from 1 (Excel row numbering, excluding header)
            result.YellowHighlightRows.Add(1); // First data row
            result.YellowHighlightRows.Add(2); // Second data row

            // Assert - Verify indices are 1-based
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[1], Is.EqualTo(2));
            Assert.That(result.YellowHighlightRows.Min(), Is.GreaterThanOrEqualTo(1));
        }

        [Test]
        public void TransformationResult_CanSetRowsProperty()
        {
            // Arrange
            var result = new TransformationResult();
            var rows = new List<TransformedRow>
            {
                new TransformedRow
                {
                    DataServizio = "2024-01-26",
                    OraInizioServizio = "09:00",
                    Assistito = "Rossi Mario",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario"
                },
                new TransformedRow
                {
                    DataServizio = "2024-01-26",
                    OraInizioServizio = "10:00",
                    Assistito = "Bianchi Anna",
                    CognomeAssistito = "Bianchi",
                    NomeAssistito = "Anna"
                }
            };

            // Act
            result.Rows = rows;

            // Assert
            Assert.That(result.Rows, Is.EqualTo(rows));
            Assert.That(result.Rows.Count, Is.EqualTo(2));
        }

        [Test]
        public void TransformationResult_CanSetYellowHighlightRowsProperty()
        {
            // Arrange
            var result = new TransformationResult();
            var highlightRows = new List<int> { 1, 3, 5, 7 };

            // Act
            result.YellowHighlightRows = highlightRows;

            // Assert
            Assert.That(result.YellowHighlightRows, Is.EqualTo(highlightRows));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(4));
        }

        [Test]
        public void TransformationResult_RequiredFields_Rows()
        {
            // Arrange
            var result = new TransformationResult
            {
                Rows = null!, // Set to null to test validation
                YellowHighlightRows = new List<int>()
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(result);
            var isValid = Validator.TryValidateObject(result, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("Rows")), Is.True);
        }

        [Test]
        public void TransformationResult_RequiredFields_YellowHighlightRows()
        {
            // Arrange
            var result = new TransformationResult
            {
                Rows = new List<TransformedRow>(),
                YellowHighlightRows = null! // Set to null to test validation
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(result);
            var isValid = Validator.TryValidateObject(result, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.False);
            Assert.That(validationResults.Any(v => v.MemberNames.Contains("YellowHighlightRows")), Is.True);
        }

        [Test]
        public void TransformationResult_AllRequiredFieldsPresent_IsValid()
        {
            // Arrange
            var result = new TransformationResult
            {
                Rows = new List<TransformedRow>(),
                YellowHighlightRows = new List<int>()
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(result);
            var isValid = Validator.TryValidateObject(result, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void TransformationResult_WithPopulatedData_IsValid()
        {
            // Arrange
            var result = new TransformationResult
            {
                Rows = new List<TransformedRow>
                {
                    new TransformedRow
                    {
                        DataServizio = "2024-01-26",
                        OraInizioServizio = "09:00",
                        Assistito = "Rossi Mario",
                        CognomeAssistito = "Rossi",
                        NomeAssistito = "Mario"
                    }
                },
                YellowHighlightRows = new List<int> { 1 }
            };

            // Act
            var validationResults = new List<ValidationResult>();
            var context = new ValidationContext(result);
            var isValid = Validator.TryValidateObject(result, context, validationResults, true);

            // Assert
            Assert.That(isValid, Is.True);
            Assert.That(validationResults.Count, Is.EqualTo(0));
        }

        [Test]
        public void TransformationResult_YellowHighlightRows_CanContainDuplicates()
        {
            // Arrange
            var result = new TransformationResult();

            // Act - Add duplicate row indices (edge case)
            result.YellowHighlightRows.Add(1);
            result.YellowHighlightRows.Add(1);
            result.YellowHighlightRows.Add(2);

            // Assert - Verify duplicates are allowed (though not expected in normal use)
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(3));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(1));
            Assert.That(result.YellowHighlightRows[1], Is.EqualTo(1));
        }

        [Test]
        public void TransformationResult_YellowHighlightRows_CanBeOutOfOrder()
        {
            // Arrange
            var result = new TransformationResult();

            // Act - Add row indices out of order
            result.YellowHighlightRows.Add(5);
            result.YellowHighlightRows.Add(2);
            result.YellowHighlightRows.Add(8);
            result.YellowHighlightRows.Add(1);

            // Assert - Verify order is preserved as added
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(4));
            Assert.That(result.YellowHighlightRows[0], Is.EqualTo(5));
            Assert.That(result.YellowHighlightRows[1], Is.EqualTo(2));
            Assert.That(result.YellowHighlightRows[2], Is.EqualTo(8));
            Assert.That(result.YellowHighlightRows[3], Is.EqualTo(1));
        }

        [Test]
        public void TransformationResult_RowsAndHighlights_CanHaveDifferentCounts()
        {
            // Arrange
            var result = new TransformationResult();
            
            // Act - Add 5 rows but only 2 highlights
            for (int i = 0; i < 5; i++)
            {
                result.Rows.Add(new TransformedRow
                {
                    DataServizio = "2024-01-26",
                    OraInizioServizio = $"{9 + i}:00",
                    Assistito = $"Person {i}",
                    CognomeAssistito = $"Last{i}",
                    NomeAssistito = $"First{i}"
                });
            }
            result.YellowHighlightRows.Add(1);
            result.YellowHighlightRows.Add(3);

            // Assert - Verify counts can differ (only some rows highlighted)
            Assert.That(result.Rows.Count, Is.EqualTo(5));
            Assert.That(result.YellowHighlightRows.Count, Is.EqualTo(2));
        }
    }
}
