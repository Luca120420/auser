using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using FsCheck;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for ExcelManager volunteer assignment methods.
    /// Tests universal properties across many generated inputs.
    /// </summary>
    [TestFixture]
    public class ExcelManagerVolunteerPropertyTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        // Feature: volunteer-email-notifications, Property 7: Volontario Column Identification
        /// <summary>
        /// Property 7: Volontario Column Identification
        /// For any Excel sheet containing a column with "Volontario" in its header (case-insensitive),
        /// the application should successfully identify that column.
        /// Validates: Requirements 4.1
        /// </summary>
        [Test]
        public void Property_IdentifiesVolontarioColumn()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generate test cases with Volontario column at different positions
            var columnPositionGen = Gen.Choose(1, 10);
            var caseVariationGen = Gen.Elements(
                "Volontario", "volontario", "VOLONTARIO", "VolOnTaRiO",
                "Volontario 1", "Volontario (Nome)", "Col. Volontario"
            );

            var testGen = Gen.Zip(columnPositionGen, caseVariationGen);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                var (columnPosition, headerText) = tuple;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    
                    // Create headers with Volontario at the specified position
                    for (int col = 1; col <= 10; col++)
                    {
                        if (col == columnPosition)
                        {
                            worksheet.Cells[1, col].Value = headerText;
                        }
                        else
                        {
                            worksheet.Cells[1, col].Value = $"Column{col}";
                        }
                    }

                    var sheet = new Sheet(worksheet);

                    // Act
                    int foundColumn = _excelManager.GetVolontarioColumnIndex(sheet);

                    // Assert
                    return foundColumn == columnPosition;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: volunteer-email-notifications, Property 9: Case-Insensitive Substring Matching
        /// <summary>
        /// Property 9: Case-Insensitive Substring Matching
        /// For any volunteer surname and any Volontario column value, if the surname appears as a substring
        /// within the column value (ignoring case), then that row should be identified as an assigned row for that volunteer.
        /// Validates: Requirements 4.3, 4.5, 4.6
        /// </summary>
        [Test]
        public void Property_CaseInsensitiveSubstringMatching()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for volunteer surnames (realistic Italian surnames)
            var surnameGen = Gen.Elements(
                "Rossi", "Bianchi", "Ferrari", "Russo", "Romano",
                "Colombo", "Ricci", "Marino", "Greco", "Bruno",
                "Gallo", "Conti", "De Luca", "Costa", "Giordano"
            );

            // Generator for case variations
            var caseVariationGen = Gen.Elements<Func<string, string>>(
                s => s.ToUpper(),           // ROSSI
                s => s.ToLower(),           // rossi
                s => s,                     // Rossi (original)
                s => char.ToUpper(s[0]) + s.Substring(1).ToLower() // Rossi
            );

            // Generator for substring contexts (how the surname appears in the Volontario column)
            var contextGen = Gen.Elements<Func<string, string>>(
                s => s,                                    // Just the surname
                s => $"Mario {s}",                        // First name + surname
                s => $"{s} Giovanni",                     // Surname + first name
                s => $"Dott. {s}",                        // Title + surname
                s => $"{s} (Coordinatore)",               // Surname + role
                s => $"Mario {s} e Luigi Verdi",          // Multiple names
                s => $"  {s}  ",                          // With whitespace
                s => $"{s}/Bianchi"                       // With separator
            );

            // Test matching cases
            var matchingTestGen = from surname in surnameGen
                                  from caseVariation in caseVariationGen
                                  from context in contextGen
                                  select (surname, caseVariation, context);

            Prop.ForAll(Arb.From(matchingTestGen), tuple =>
            {
                string surname = tuple.Item1;
                Func<string, string> caseVariation = tuple.Item2;
                Func<string, string> context = tuple.Item3;
                var volontarioValue = context(caseVariation(surname));

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    
                    // Create header row
                    worksheet.Cells[1, 1].Value = "Volontario";
                    worksheet.Cells[1, 2].Value = "Data";
                    worksheet.Cells[1, 3].Value = "Servizio";

                    // Create a data row with the Volontario value
                    worksheet.Cells[2, 1].Value = volontarioValue;
                    worksheet.Cells[2, 2].Value = "2024-01-15";
                    worksheet.Cells[2, 3].Value = "Trasporto";

                    var sheet = new Sheet(worksheet);
                    var volunteers = new Dictionary<string, string>
                    {
                        { surname, $"{surname.ToLower()}@example.com" }
                    };

                    // Act
                    var assignments = _excelManager.IdentifyVolunteerAssignments(sheet, volunteers);

                    // Assert - should find exactly one assignment with one row
                    if (assignments.Count != 1)
                    {
                        throw new Exception($"Expected 1 assignment, got {assignments.Count}. Surname: '{surname}', Volontario value: '{volontarioValue}'");
                    }
                    if (assignments[0].Surname != surname)
                    {
                        throw new Exception($"Expected surname '{surname}', got '{assignments[0].Surname}'");
                    }
                    if (assignments[0].AssignedRows.Count != 1)
                    {
                        throw new Exception($"Expected 1 assigned row, got {assignments[0].AssignedRows.Count}");
                    }
                    // Note: Excel/EPPlus may trim whitespace, so we compare the actual stored value
                    var actualVolontarioValue = assignments[0].AssignedRows[0]["Volontario"];
                    if (actualVolontarioValue != volontarioValue.Trim())
                    {
                        throw new Exception($"Expected Volontario value '{volontarioValue.Trim()}', got '{actualVolontarioValue}'");
                    }
                    
                    return true;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: volunteer-email-notifications, Property 9: Case-Insensitive Substring Matching (Non-Matching)
        /// <summary>
        /// Property 9: Case-Insensitive Substring Matching - Non-Matching Cases
        /// For any volunteer surname and any Volontario column value that does NOT contain the surname,
        /// that row should NOT be identified as an assigned row for that volunteer.
        /// Validates: Requirements 4.3, 4.5, 4.6
        /// </summary>
        [Test]
        public void Property_CaseInsensitiveSubstringMatching_NonMatching()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for volunteer surnames
            var surnameGen = Gen.Elements(
                "Rossi", "Bianchi", "Ferrari", "Russo", "Romano"
            );

            // Generator for non-matching Volontario values
            var nonMatchGen = Gen.Elements(
                "Verdi", "Esposito", "Fontana", "Caruso", "Moretti",
                "", "   ", "N/A", "TBD", "Mario Verdi", "Dott. Esposito"
            );

            var testGen = from surname in surnameGen
                          from volontarioValue in nonMatchGen
                          select (surname, volontarioValue);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                string surname = tuple.Item1;
                string volontarioValue = tuple.Item2;

                // Skip if the non-match value actually contains the surname (edge case)
                if (!string.IsNullOrWhiteSpace(volontarioValue) &&
                    volontarioValue.IndexOf(surname, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true; // Skip this test case
                }

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    
                    // Create header row
                    worksheet.Cells[1, 1].Value = "Volontario";
                    worksheet.Cells[1, 2].Value = "Data";

                    // Create a data row with the non-matching Volontario value
                    worksheet.Cells[2, 1].Value = volontarioValue;
                    worksheet.Cells[2, 2].Value = "2024-01-15";

                    var sheet = new Sheet(worksheet);
                    var volunteers = new Dictionary<string, string>
                    {
                        { surname, $"{surname.ToLower()}@example.com" }
                    };

                    // Act
                    var assignments = _excelManager.IdentifyVolunteerAssignments(sheet, volunteers);

                    // Assert - should find no assignments
                    return assignments.Count == 0;
                }
            }).QuickCheckThrowOnFailure();
        }

        // Feature: volunteer-email-notifications, Property 9: Case-Insensitive Substring Matching (Multiple Rows)
        /// <summary>
        /// Property 9: Case-Insensitive Substring Matching - Multiple Rows
        /// For any volunteer surname, if N rows contain the surname in the Volontario column,
        /// then exactly N rows should be identified as assigned rows for that volunteer.
        /// Validates: Requirements 4.3, 4.5, 4.6
        /// </summary>
        [Test]
        public void Property_CaseInsensitiveSubstringMatching_MultipleRows()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generator for volunteer surnames
            var surnameGen = Gen.Elements(
                "Rossi", "Bianchi", "Ferrari", "Russo", "Romano"
            );

            // Generator for number of matching rows (1-10)
            var rowCountGen = Gen.Choose(1, 10);

            // Generator for case variations
            var caseVariationGen = Gen.Elements<Func<string, string>>(
                s => s.ToUpper(),
                s => s.ToLower(),
                s => s
            );

            var testGen = from surname in surnameGen
                          from matchingRowCount in rowCountGen
                          from caseVariation in caseVariationGen
                          select (surname, matchingRowCount, caseVariation);

            Prop.ForAll(Arb.From(testGen), tuple =>
            {
                string surname = tuple.Item1;
                int matchingRowCount = tuple.Item2;
                Func<string, string> caseVariation = tuple.Item3;

                using (var package = new ExcelPackage())
                {
                    var worksheet = package.Workbook.Worksheets.Add("Test");
                    
                    // Create header row
                    worksheet.Cells[1, 1].Value = "Volontario";
                    worksheet.Cells[1, 2].Value = "Data";
                    worksheet.Cells[1, 3].Value = "Servizio";

                    // Create matching rows
                    for (int i = 0; i < matchingRowCount; i++)
                    {
                        int rowNum = i + 2;
                        worksheet.Cells[rowNum, 1].Value = $"Mario {caseVariation(surname)}";
                        worksheet.Cells[rowNum, 2].Value = $"2024-01-{15 + i}";
                        worksheet.Cells[rowNum, 3].Value = $"Servizio {i + 1}";
                    }

                    // Add some non-matching rows
                    for (int i = 0; i < 3; i++)
                    {
                        int rowNum = matchingRowCount + 2 + i;
                        worksheet.Cells[rowNum, 1].Value = "Verdi";
                        worksheet.Cells[rowNum, 2].Value = $"2024-01-{20 + i}";
                        worksheet.Cells[rowNum, 3].Value = $"Servizio {matchingRowCount + i + 1}";
                    }

                    var sheet = new Sheet(worksheet);
                    var volunteers = new Dictionary<string, string>
                    {
                        { surname, $"{surname.ToLower()}@example.com" }
                    };

                    // Act
                    var assignments = _excelManager.IdentifyVolunteerAssignments(sheet, volunteers);

                    // Assert - should find exactly one assignment with matchingRowCount rows
                    return assignments.Count == 1 &&
                           assignments[0].Surname == surname &&
                           assignments[0].AssignedRows.Count == matchingRowCount;
                }
            }).QuickCheckThrowOnFailure();
        }
    }
}
