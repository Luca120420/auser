using System;
using System.Collections.Generic;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for ColumnStructureManager class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 2.1, 3.1
    /// </summary>
    [TestFixture]
    public class ColumnStructureManagerPropertyTests
    {
        private IColumnStructureManager _columnStructureManager = null!;

        [SetUp]
        public void Setup()
        {
            _columnStructureManager = new ColumnStructureManager();
        }

        // Feature: excel-output-enhancement, Property 1: Column Positioning Invariant
        /// <summary>
        /// Property 1: Column Positioning Invariant
        /// For any generated Excel output, the "Indirizzo" column index SHALL equal
        /// the "Assistito" column index plus one.
        /// **Validates: Requirements 2.1**
        /// </summary>
        [Test]
        public void Property_ColumnPositioningInvariant()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.Default.Unit(),
                (_) =>
                {
                    try
                    {
                        // Act - Get column headers
                        var headers = _columnStructureManager.GetColumnHeaders();

                        // Get indices for Assistito and Indirizzo columns
                        int assistitoIndex = _columnStructureManager.GetColumnIndex("Assistito");
                        int indirizzoIndex = _columnStructureManager.GetColumnIndex("Indirizzo");

                        // Assert - Indirizzo should be immediately after Assistito
                        var positioningCorrect = indirizzoIndex == assistitoIndex + 1;

                        // Additional assertions for robustness
                        var assistitoExists = assistitoIndex >= 0;
                        var indirizzoExists = indirizzoIndex >= 0;
                        var bothColumnsPresent = assistitoExists && indirizzoExists;

                        // Verify the columns are actually in the header list
                        var assistitoInHeaders = headers.Contains("Assistito");
                        var indirizzoInHeaders = headers.Contains("Indirizzo");

                        if (!bothColumnsPresent)
                        {
                            return false.Label($"Required columns not found: Assistito index={assistitoIndex}, Indirizzo index={indirizzoIndex}");
                        }

                        if (!assistitoInHeaders || !indirizzoInHeaders)
                        {
                            return false.Label($"Required columns not in headers: Assistito={assistitoInHeaders}, Indirizzo={indirizzoInHeaders}");
                        }

                        if (!positioningCorrect)
                        {
                            return false.Label($"Column positioning invariant violated: Assistito at index {assistitoIndex}, Indirizzo at index {indirizzoIndex} (expected {assistitoIndex + 1})");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Column positioning check failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: excel-output-enhancement, Property 2: Column Exclusion
        /// <summary>
        /// Property 2: Column Exclusion
        /// For any generated Excel output, the column headers SHALL NOT contain "Comune Partenza".
        /// **Validates: Requirements 3.1**
        /// </summary>
        [Test]
        public void Property_ColumnExclusion()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.Default.Unit(),
                (_) =>
                {
                    try
                    {
                        // Act - Get column headers
                        var headers = _columnStructureManager.GetColumnHeaders();

                        // Assert - "Comune Partenza" should NOT be in the headers
                        var comunePartenzaPresent = headers.Contains("Comune Partenza");

                        if (comunePartenzaPresent)
                        {
                            return false.Label($"Column exclusion violated: 'Comune Partenza' found in headers at index {headers.IndexOf("Comune Partenza")}");
                        }

                        // Additional verification: ensure we have a valid header list
                        if (headers == null || headers.Count == 0)
                        {
                            return false.Label("Column headers list is null or empty");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Column exclusion check failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }
    }
}
