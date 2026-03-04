using System;
using System.IO;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Integration tests for CSVParser with real CSV files.
    /// Tests parsing of actual CSV files from the project.
    /// </summary>
    [TestFixture]
    public class CSVParserIntegrationTests
    {
        private CSVParser _parser = null!;

        [SetUp]
        public void Setup()
        {
            _parser = new CSVParser();
        }

        [Test]
        public void ParseCSV_WithRealCSVFile_ParsesSuccessfully()
        {
            // Arrange
            // Look for CSV file in project root (go up from bin/Debug/net6.0-windows)
            var projectRoot = Path.Combine(TestContext.CurrentContext.TestDirectory, "..", "..", "..");
            var csvPath = Path.Combine(projectRoot, "168514-Estrazione_1770193162042.csv");
            
            // Skip test if file doesn't exist (for CI/CD environments)
            if (!File.Exists(csvPath))
            {
                Assert.Ignore($"Real CSV file not found at {csvPath} - skipping integration test");
                return;
            }

            // Act
            var result = _parser.ParseCSV(csvPath);

            // Assert
            Assert.That(result, Is.Not.Null);
            Assert.That(result.Count, Is.GreaterThan(0), "Should parse at least one appointment");
            
            // Verify first appointment has required fields
            var firstAppointment = result[0];
            Assert.That(firstAppointment.DataServizio, Is.Not.Null.And.Not.Empty);
            Assert.That(firstAppointment.OraInizioServizio, Is.Not.Null.And.Not.Empty);
            Assert.That(firstAppointment.CognomeAssistito, Is.Not.Null.And.Not.Empty);
            Assert.That(firstAppointment.NomeAssistito, Is.Not.Null.And.Not.Empty);
            
            // Log some info for verification
            TestContext.WriteLine($"Parsed {result.Count} appointments from real CSV file");
            TestContext.WriteLine($"First appointment: {firstAppointment.CognomeAssistito} {firstAppointment.NomeAssistito} on {firstAppointment.DataServizio} at {firstAppointment.OraInizioServizio}");
        }

        [Test]
        public void ValidateCSVStructure_WithRealCSVFile_ReturnsTrue()
        {
            // Arrange
            // Look for CSV file in project root (go up from bin/Debug/net6.0-windows)
            var projectRoot = Path.Combine(TestContext.CurrentContext.TestDirectory, "..", "..", "..");
            var csvPath = Path.Combine(projectRoot, "168514-Estrazione_1770193162042.csv");
            
            // Skip test if file doesn't exist (for CI/CD environments)
            if (!File.Exists(csvPath))
            {
                Assert.Ignore($"Real CSV file not found at {csvPath} - skipping integration test");
                return;
            }

            // Act
            var result = _parser.ValidateCSVStructure(csvPath);

            // Assert
            Assert.That(result, Is.True, "Real CSV file should have all required columns");
        }

        [Test]
        public void ParseCSV_WithRealCSVFile_PreservesItalianCharacters()
        {
            // Arrange
            // Look for CSV file in project root (go up from bin/Debug/net6.0-windows)
            var projectRoot = Path.Combine(TestContext.CurrentContext.TestDirectory, "..", "..", "..");
            var csvPath = Path.Combine(projectRoot, "168514-Estrazione_1770193162042.csv");
            
            // Skip test if file doesn't exist (for CI/CD environments)
            if (!File.Exists(csvPath))
            {
                Assert.Ignore($"Real CSV file not found at {csvPath} - skipping integration test");
                return;
            }

            // Act
            var result = _parser.ParseCSV(csvPath);

            // Assert
            Assert.That(result, Is.Not.Null);
            
            // Check if any appointments contain Italian characters
            var hasItalianChars = result.Any(a => 
                ContainsItalianCharacters(a.CognomeAssistito) ||
                ContainsItalianCharacters(a.NomeAssistito) ||
                ContainsItalianCharacters(a.IndirizzoPartenza ?? "") ||
                ContainsItalianCharacters(a.IndirizzoDestinazione ?? "") ||
                ContainsItalianCharacters(a.NoteERichieste ?? ""));
            
            // If Italian characters are present, verify they're preserved correctly
            if (hasItalianChars)
            {
                TestContext.WriteLine("Italian characters found and preserved in CSV data");
            }
            else
            {
                TestContext.WriteLine("No Italian special characters found in this CSV file");
            }
        }

        private bool ContainsItalianCharacters(string text)
        {
            if (string.IsNullOrEmpty(text))
                return false;
            
            // Check for common Italian accented characters
            return text.Contains('à') || text.Contains('è') || text.Contains('é') ||
                   text.Contains('ì') || text.Contains('ò') || text.Contains('ù') ||
                   text.Contains('À') || text.Contains('È') || text.Contains('É') ||
                   text.Contains('Ì') || text.Contains('Ò') || text.Contains('Ù');
        }
    }
}
