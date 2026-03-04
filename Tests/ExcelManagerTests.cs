using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using OfficeOpenXml;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for ExcelManager class.
    /// Tests Requirements 3.1, 3.2, 3.4 - Excel workbook management operations.
    /// </summary>
    [TestFixture]
    public class ExcelManagerTests
    {
        private ExcelManager? _excelManager;
        private string? _testFilePath;

        [SetUp]
        public void Setup()
        {
            _excelManager = new ExcelManager();
            _testFilePath = Path.Combine(Path.GetTempPath(), $"test_{Guid.NewGuid()}.xlsx");
            
            // Set EPPlus license context
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        [TearDown]
        public void TearDown()
        {
            // Clean up test files
            if (_testFilePath != null && File.Exists(_testFilePath))
            {
                File.Delete(_testFilePath);
            }
        }

        #region OpenWorkbook Tests

        [Test]
        public void OpenWorkbook_WithValidFile_ReturnsExcelWorkbook()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);

            // Act
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Assert
            Assert.That(workbook, Is.Not.Null);
            Assert.That(workbook.Package, Is.Not.Null);
        }

        [Test]
        public void OpenWorkbook_WithNonExistentFile_ThrowsFileNotFoundException()
        {
            // Arrange
            var nonExistentPath = "nonexistent.xlsx";

            // Act & Assert
            var ex = Assert.Throws<FileNotFoundException>(() => _excelManager!.OpenWorkbook(nonExistentPath));
            Assert.That(ex!.Message, Does.Contain("Il file Excel non è stato trovato"));
        }

        #endregion

        #region GetSheetNames Tests

        [Test]
        public void GetSheetNames_WithMultipleSheets_ReturnsAllSheetNames()
        {
            // Arrange
            CreateTestExcelFileWithSheets(_testFilePath!, new[] { "1", "2", "fissi" });
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act
            var sheetNames = _excelManager.GetSheetNames(workbook);

            // Assert
            Assert.That(sheetNames.Count, Is.EqualTo(3));
            Assert.That(sheetNames, Does.Contain("1"));
            Assert.That(sheetNames, Does.Contain("2"));
            Assert.That(sheetNames, Does.Contain("fissi"));
        }

        [Test]
        public void GetSheetNames_WithSingleSheet_ReturnsSheetName()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act
            var sheetNames = _excelManager.GetSheetNames(workbook);

            // Assert
            Assert.That(sheetNames.Count, Is.EqualTo(1));
            Assert.That(sheetNames, Does.Contain("Sheet1"));
        }

        [Test]
        public void GetSheetNames_WithNullWorkbook_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.GetSheetNames(null!));
        }

        #endregion

        #region GetNextSheetNumber Tests

        [Test]
        public void GetNextSheetNumber_WithNumberedSheets_ReturnsMaxPlusOne()
        {
            // Arrange
            var sheetNames = new List<string> { "1", "2", "3", "fissi" };

            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(sheetNames);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(4));
        }

        [Test]
        public void GetNextSheetNumber_WithNonSequentialNumbers_ReturnsMaxPlusOne()
        {
            // Arrange
            var sheetNames = new List<string> { "1", "5", "3", "fissi" };

            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(sheetNames);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(6));
        }

        [Test]
        public void GetNextSheetNumber_WithNoNumberedSheets_ReturnsOne()
        {
            // Arrange
            var sheetNames = new List<string> { "fissi", "other" };

            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(sheetNames);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(1));
        }

        [Test]
        public void GetNextSheetNumber_WithEmptyList_ReturnsOne()
        {
            // Arrange
            var sheetNames = new List<string>();

            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(sheetNames);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(1));
        }

        [Test]
        public void GetNextSheetNumber_WithNullList_ReturnsOne()
        {
            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(null!);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(1));
        }

        [Test]
        public void GetNextSheetNumber_WithMixedSheetNames_IgnoresNonNumeric()
        {
            // Arrange
            var sheetNames = new List<string> { "1", "2", "fissi", "Sheet1", "10" };

            // Act
            var nextNumber = _excelManager!.GetNextSheetNumber(sheetNames);

            // Assert
            Assert.That(nextNumber, Is.EqualTo(11));
        }

        #endregion

        #region GetFissiSheet Tests

        [Test]
        public void GetFissiSheet_WithFissiSheet_ReturnsSheet()
        {
            // Arrange
            CreateTestExcelFileWithSheets(_testFilePath!, new[] { "1", "fissi" });
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act
            var fissiSheet = _excelManager.GetFissiSheet(workbook);

            // Assert
            Assert.That(fissiSheet, Is.Not.Null);
            Assert.That(fissiSheet.Name, Is.EqualTo("fissi"));
        }

        [Test]
        public void GetFissiSheet_CaseInsensitive_ReturnsSheet()
        {
            // Arrange
            CreateTestExcelFileWithSheets(_testFilePath!, new[] { "1", "FISSI" });
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act
            var fissiSheet = _excelManager.GetFissiSheet(workbook);

            // Assert
            Assert.That(fissiSheet, Is.Not.Null);
            Assert.That(fissiSheet.Name, Is.EqualTo("FISSI"));
        }

        [Test]
        public void GetFissiSheet_WithoutFissiSheet_ThrowsInvalidOperationException()
        {
            // Arrange
            CreateTestExcelFileWithSheets(_testFilePath!, new[] { "1", "2" });
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act & Assert
            var ex = Assert.Throws<InvalidOperationException>(() => _excelManager.GetFissiSheet(workbook));
            Assert.That(ex!.Message, Does.Contain("Il foglio 'fissi' non è stato trovato"));
        }

        [Test]
        public void GetFissiSheet_WithNullWorkbook_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.GetFissiSheet(null!));
        }

        #endregion

        #region CreateNewSheet Tests

        [Test]
        public void CreateNewSheet_CreatesSheetWithCorrectName()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);

            // Act
            var newSheet = _excelManager.CreateNewSheet(workbook, 5);

            // Assert
            Assert.That(newSheet, Is.Not.Null);
            Assert.That(newSheet.Name, Is.EqualTo("5"));
        }

        [Test]
        public void CreateNewSheet_WithNullWorkbook_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.CreateNewSheet(null!, 1));
        }

        #endregion

        #region WriteHeader Tests

        [Test]
        public void WriteHeader_WritesTextToFirstRow()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);
            var sheet = _excelManager.CreateNewSheet(workbook, 1);
            var headerText = "26 gen 01 feb Settimana 5referente settimana = Test";

            // Act
            _excelManager.WriteHeader(sheet, headerText);

            // Assert
            var cellValue = sheet.Worksheet.Cells[1, 1].Value?.ToString();
            Assert.That(cellValue, Is.EqualTo(headerText));
        }

        [Test]
        public void WriteHeader_WithNullSheet_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.WriteHeader(null!, "test"));
        }

        #endregion

        #region WriteDataRows Tests

        [Test]
        public void WriteDataRows_WritesAllRowsCorrectly()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);
            var sheet = _excelManager.CreateNewSheet(workbook, 1);
            
            var rows = new List<TransformedRow>
            {
                new TransformedRow
                {
                    DataServizio = "2024-01-26",
                    OraInizioServizio = "09:00",
                    Assistito = "Rossi Mario",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario",
                    Indirizzo = "Milano Via Roma 1",
                    Destinazione = "Ospedale",
                    EmptyColumn1 = "",
                    EmptyColumn2 = "",
                    EmptyColumn3 = "",
                    EmptyColumn4 = "",
                    EmptyColumn5 = "",
                    OraInizioServizioCopy = "09:00",
                    Partenza = "",
                    NoteERichieste = "Test note"
                }
            };

            // Act
            _excelManager.WriteDataRows(sheet, rows, 2);

            // Assert
            Assert.That(sheet.Worksheet.Cells[2, 1].Value?.ToString(), Is.EqualTo("2024-01-26"));
            Assert.That(sheet.Worksheet.Cells[2, 2].Value?.ToString(), Is.EqualTo("09:00"));
            Assert.That(sheet.Worksheet.Cells[2, 3].Value?.ToString(), Is.EqualTo("Rossi Mario"));
            Assert.That(sheet.Worksheet.Cells[2, 4].Value?.ToString(), Is.EqualTo("Rossi"));
            Assert.That(sheet.Worksheet.Cells[2, 5].Value?.ToString(), Is.EqualTo("Mario"));
            Assert.That(sheet.Worksheet.Cells[2, 6].Value?.ToString(), Is.EqualTo("Milano Via Roma 1"));
            Assert.That(sheet.Worksheet.Cells[2, 7].Value?.ToString(), Is.EqualTo("Ospedale"));
            Assert.That(sheet.Worksheet.Cells[2, 13].Value?.ToString(), Is.EqualTo("09:00"));
            Assert.That(sheet.Worksheet.Cells[2, 15].Value?.ToString(), Is.EqualTo("Test note"));
        }

        [Test]
        public void WriteDataRows_WithNullSheet_ThrowsArgumentNullException()
        {
            // Arrange
            var rows = new List<TransformedRow>();

            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.WriteDataRows(null!, rows, 2));
        }

        [Test]
        public void WriteDataRows_WithNullRows_DoesNotThrow()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);
            var sheet = _excelManager.CreateNewSheet(workbook, 1);

            // Act & Assert
            Assert.DoesNotThrow(() => _excelManager.WriteDataRows(sheet, null!, 2));
        }

        #endregion

        #region SaveWorkbook Tests

        [Test]
        public void SaveWorkbook_SavesFileSuccessfully()
        {
            // Arrange
            CreateTestExcelFile(_testFilePath!);
            var workbook = _excelManager!.OpenWorkbook(_testFilePath!);
            var outputPath = Path.Combine(Path.GetTempPath(), $"output_{Guid.NewGuid()}.xlsx");

            try
            {
                // Act
                _excelManager.SaveWorkbook(workbook, outputPath);

                // Assert
                Assert.That(File.Exists(outputPath), Is.True);
            }
            finally
            {
                if (File.Exists(outputPath))
                {
                    File.Delete(outputPath);
                }
            }
        }

        [Test]
        public void SaveWorkbook_WithNullWorkbook_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<ArgumentNullException>(() => _excelManager!.SaveWorkbook(null!, "test.xlsx"));
        }

        #endregion

        #region Helper Methods

        private void CreateTestExcelFile(string filePath)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // EPPlus requires at least one worksheet
                package.Workbook.Worksheets.Add("Sheet1");
                package.Save();
            }
        }

        private void CreateTestExcelFileWithSheets(string filePath, string[] sheetNames)
        {
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                foreach (var sheetName in sheetNames)
                {
                    package.Workbook.Worksheets.Add(sheetName);
                }
                package.Save();
            }
        }

        #endregion
    }
}
