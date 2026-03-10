using System;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using NUnit.Framework;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Bug condition exploration tests for fissi time format fix.
    /// 
    /// CRITICAL: These tests are EXPECTED TO FAIL on unfixed code.
    /// Test failure confirms the bug exists.
    /// 
    /// These tests encode the EXPECTED (correct) behavior.
    /// When the bug is fixed, these tests will pass.
    /// 
    /// **Validates: Requirements 1.1, 1.2, 2.1, 2.2, 2.3**
    /// </summary>
    [TestFixture]
    public class FissiTimeFormatBugExplorationTests
    {
        private ExcelManager _excelManager;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            _excelManager = new ExcelManager();
        }

        /// <summary>
        /// Property 1: Fault Condition - Time Format Display for Partenza (Column 2)
        /// 
        /// This test creates a fissi sheet with decimal time values in column 2 (Partenza)
        /// and verifies that the target cells display in time format (h:mm) instead of decimal format.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL because:
        /// - Target cells in column 2 display decimal values (e.g., "0.354166666666667") instead of time format (e.g., "8:30")
        /// - The number format is not set to "h:mm" or the format is not applied correctly
        /// 
        /// When the bug is fixed, this test will PASS.
        /// 
        /// **Validates: Requirements 1.1, 2.1, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_PartenzaColumn_DecimalTimeValue_ShouldDisplayAsTimeFormat()
        {
            // Arrange - Create a fissi sheet with decimal time value in column 2
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";
                fissiWorksheet.Cells[2, 3].Value = "Assistito";

                // Add data row with decimal time value in column 2
                // 0.354166666666667 represents 8:30 AM in Excel time format
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 2].Value = 0.354166666666667;  // Decimal time value
                fissiWorksheet.Cells[3, 3].Value = "Rossi Mario";

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act - Call AppendFissiData on UNFIXED code
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify CORRECT behavior (will fail on unfixed code)
                var targetCell = targetWorksheet.Cells[1, 2];
                
                // Requirement 2.3: Number format should be "h:mm"
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                    "COUNTEREXAMPLE: Target cell in column 2 (Partenza) should have number format 'h:mm'. " +
                    $"Expected 'h:mm', but found '{targetCell.Style.Numberformat.Format}'. " +
                    "This confirms the bug: time format is not being applied to Partenza column.");

                // Requirement 2.1: Cell text should display as time (e.g., "8:30")
                var cellText = targetCell.Text;
                Assert.That(cellText, Does.Match(@"^\d{1,2}:\d{2}$"),
                    "COUNTEREXAMPLE: Target cell in column 2 (Partenza) should display as time format (e.g., '8:30'). " +
                    $"Expected time pattern (h:mm), but found '{cellText}'. " +
                    "This confirms the bug: decimal value is displayed instead of time format.");
            }
        }

        /// <summary>
        /// Property 1: Fault Condition - Time Format Display for Arrivo (Column 9)
        /// 
        /// This test creates a fissi sheet with decimal time values in column 9 (Arrivo)
        /// and verifies that the target cells display in time format (h:mm) instead of decimal format.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL because:
        /// - Target cells in column 9 display decimal values (e.g., "0.416666666666667") instead of time format (e.g., "10:00")
        /// - The number format is not set to "h:mm" or the format is not applied correctly
        /// 
        /// When the bug is fixed, this test will PASS.
        /// 
        /// **Validates: Requirements 1.2, 2.2, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_ArrivoColumn_DecimalTimeValue_ShouldDisplayAsTimeFormat()
        {
            // Arrange - Create a fissi sheet with decimal time value in column 9
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                // Add data row with decimal time value in column 9
                // 0.416666666666667 represents 10:00 AM in Excel time format
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 9].Value = 0.416666666666667;  // Decimal time value

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act - Call AppendFissiData on UNFIXED code
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Verify CORRECT behavior (will fail on unfixed code)
                var targetCell = targetWorksheet.Cells[1, 9];
                
                // Requirement 2.3: Number format should be "h:mm"
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                    "COUNTEREXAMPLE: Target cell in column 9 (Arrivo) should have number format 'h:mm'. " +
                    $"Expected 'h:mm', but found '{targetCell.Style.Numberformat.Format}'. " +
                    "This confirms the bug: time format is not being applied to Arrivo column.");

                // Requirement 2.2: Cell text should display as time (e.g., "10:00")
                var cellText = targetCell.Text;
                Assert.That(cellText, Does.Match(@"^\d{1,2}:\d{2}$"),
                    "COUNTEREXAMPLE: Target cell in column 9 (Arrivo) should display as time format (e.g., '10:00'). " +
                    $"Expected time pattern (h:mm), but found '{cellText}'. " +
                    "This confirms the bug: decimal value is displayed instead of time format.");
            }
        }

        /// <summary>
        /// Property 1: Fault Condition - Edge Case Midnight (0.0)
        /// 
        /// Tests that midnight (0.0) displays correctly as "0:00" in time format.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL.
        /// 
        /// **Validates: Requirements 2.1, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_PartenzaColumn_Midnight_ShouldDisplayAsTimeFormat()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";

                // 0.0 represents midnight (00:00)
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 2].Value = 0.0;

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert
                var targetCell = targetWorksheet.Cells[1, 2];
                
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                    "COUNTEREXAMPLE: Midnight value should have 'h:mm' format. " +
                    $"Found '{targetCell.Style.Numberformat.Format}'.");

                var cellText = targetCell.Text;
                Assert.That(cellText, Does.Match(@"^\d{1,2}:\d{2}$"),
                    "COUNTEREXAMPLE: Midnight should display as time format (e.g., '0:00'). " +
                    $"Found '{cellText}'.");
            }
        }

        /// <summary>
        /// Property 1: Fault Condition - Edge Case Noon (0.5)
        /// 
        /// Tests that noon (0.5) displays correctly as "12:00" in time format.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL.
        /// 
        /// **Validates: Requirements 2.2, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_ArrivoColumn_Noon_ShouldDisplayAsTimeFormat()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                // 0.5 represents noon (12:00)
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 9].Value = 0.5;

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert
                var targetCell = targetWorksheet.Cells[1, 9];
                
                Assert.That(targetCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                    "COUNTEREXAMPLE: Noon value should have 'h:mm' format. " +
                    $"Found '{targetCell.Style.Numberformat.Format}'.");

                var cellText = targetCell.Text;
                Assert.That(cellText, Does.Match(@"^\d{1,2}:\d{2}$"),
                    "COUNTEREXAMPLE: Noon should display as time format (e.g., '12:00'). " +
                    $"Found '{cellText}'.");
            }
        }

        /// <summary>
        /// Property 1: Fault Condition - Multiple Rows with Different Time Values
        /// 
        /// Tests that multiple rows with different decimal time values all display correctly in time format.
        /// 
        /// EXPECTED OUTCOME ON UNFIXED CODE: This test will FAIL for all rows.
        /// 
        /// **Validates: Requirements 1.1, 1.2, 2.1, 2.2, 2.3**
        /// </summary>
        [Test]
        public void BugExploration_MultipleRows_DecimalTimeValues_ShouldDisplayAsTimeFormat()
        {
            // Arrange
            using (var package = new ExcelPackage())
            {
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                var targetWorksheet = package.Workbook.Worksheets.Add("target");

                // Add header row
                fissiWorksheet.Cells[2, 1].Value = "Data";
                fissiWorksheet.Cells[2, 2].Value = "Partenza";
                fissiWorksheet.Cells[2, 9].Value = "Arrivo";

                // Add multiple data rows with different time values
                // Row 1: 8:30 and 10:00
                fissiWorksheet.Cells[3, 1].Value = "15/01/2024";
                fissiWorksheet.Cells[3, 2].Value = 0.354166666666667;  // 8:30
                fissiWorksheet.Cells[3, 9].Value = 0.416666666666667;  // 10:00

                // Row 2: 14:00 and 16:30
                fissiWorksheet.Cells[4, 1].Value = "16/01/2024";
                fissiWorksheet.Cells[4, 2].Value = 0.583333333333333;  // 14:00
                fissiWorksheet.Cells[4, 9].Value = 0.6875;             // 16:30

                // Row 3: 9:15 and 11:45
                fissiWorksheet.Cells[5, 1].Value = "17/01/2024";
                fissiWorksheet.Cells[5, 2].Value = 0.385416666666667;  // 9:15
                fissiWorksheet.Cells[5, 9].Value = 0.489583333333333;  // 11:45

                var fissiSheet = new Sheet(fissiWorksheet);
                var targetSheet = new Sheet(targetWorksheet);

                // Act
                _excelManager.AppendFissiData(targetSheet, fissiSheet, 1, DateTime.Now);

                // Assert - Check all rows
                for (int row = 1; row <= 3; row++)
                {
                    // Check Partenza column (2)
                    var partenzaCell = targetWorksheet.Cells[row, 2];
                    Assert.That(partenzaCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                        $"COUNTEREXAMPLE Row {row}: Partenza should have 'h:mm' format. " +
                        $"Found '{partenzaCell.Style.Numberformat.Format}'.");
                    
                    Assert.That(partenzaCell.Text, Does.Match(@"^\d{1,2}:\d{2}$"),
                        $"COUNTEREXAMPLE Row {row}: Partenza should display as time. " +
                        $"Found '{partenzaCell.Text}'.");

                    // Check Arrivo column (9)
                    var arrivoCell = targetWorksheet.Cells[row, 9];
                    Assert.That(arrivoCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                        $"COUNTEREXAMPLE Row {row}: Arrivo should have 'h:mm' format. " +
                        $"Found '{arrivoCell.Style.Numberformat.Format}'.");
                    
                    Assert.That(arrivoCell.Text, Does.Match(@"^\d{1,2}:\d{2}$"),
                        $"COUNTEREXAMPLE Row {row}: Arrivo should display as time. " +
                        $"Found '{arrivoCell.Text}'.");
                }
            }
        }
    }
}
