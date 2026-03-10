using System;
using System.Collections.Generic;
using System.IO;
using NUnit.Framework;
using Moq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Properties;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for ApplicationController class.
    /// Tests workflow orchestration, error handling, and state management.
    /// Validates: Requirements 7.1, 7.3, 7.4, 7.5, 7.6, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6
    /// </summary>
    [TestFixture]
    public class ApplicationControllerTests
    {
        private Mock<IGUI> _mockGui = null!;
        private Mock<ICSVParser> _mockCsvParser = null!;
        private Mock<IExcelManager> _mockExcelManager = null!;
        private Mock<IDataTransformer> _mockDataTransformer = null!;
        private Mock<IHeaderCalculator> _mockHeaderCalculator = null!;
        private ApplicationController _controller = null!;

        [SetUp]
        public void Setup()
        {
            _mockGui = new Mock<IGUI>();
            _mockCsvParser = new Mock<ICSVParser>();
            _mockExcelManager = new Mock<IExcelManager>();
            _mockDataTransformer = new Mock<IDataTransformer>();
            _mockHeaderCalculator = new Mock<IHeaderCalculator>();

            _controller = new ApplicationController(
                _mockGui.Object,
                _mockCsvParser.Object,
                _mockExcelManager.Object,
                _mockDataTransformer.Object,
                _mockHeaderCalculator.Object
            );
        }

        #region OnCSVFileSelected Tests

        [Test]
        public void OnCSVFileSelected_WithValidFile_StoresPathAndEnablesProcessButton()
        {
            // Arrange
            string csvPath = Path.GetTempFileName();
            File.WriteAllText(csvPath, "test");
            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out It.Ref<List<string>>.IsAny))
                .Returns(true);

            // Act
            _controller.OnCSVFileSelected(csvPath);

            // Assert
            _mockGui.Verify(g => g.DisplaySelectedCSVPath(csvPath), Times.Once);
            Assert.That(_controller.CanProcess, Is.False); // Excel not selected yet

            // Cleanup
            File.Delete(csvPath);
        }

        [Test]
        public void OnCSVFileSelected_WithNonExistentFile_ShowsError()
        {
            // Arrange
            string csvPath = "nonexistent.csv";
            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out It.Ref<List<string>>.IsAny))
                .Throws<FileNotFoundException>();

            // Act
            _controller.OnCSVFileSelected(csvPath);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorCSVFileRead), Times.Once);
            _mockGui.Verify(g => g.DisplaySelectedCSVPath(It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnCSVFileSelected_WithMissingColumns_ShowsErrorWithColumnNames()
        {
            // Arrange
            string csvPath = Path.GetTempFileName();
            File.WriteAllText(csvPath, "test");
            var missingColumns = new List<string> { "DATA SERVIZIO", "ORA INIZIO SERVIZIO" };
            
            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out missingColumns))
                .Returns(false);

            // Act
            _controller.OnCSVFileSelected(csvPath);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(It.Is<string>(msg => 
                msg.Contains("DATA SERVIZIO") && msg.Contains("ORA INIZIO SERVIZIO"))), Times.Once);

            // Cleanup
            File.Delete(csvPath);
        }

        [Test]
        public void OnCSVFileSelected_WithNullPath_DoesNothing()
        {
            // Act
            _controller.OnCSVFileSelected(null);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(It.IsAny<string>()), Times.Never);
            _mockGui.Verify(g => g.DisplaySelectedCSVPath(It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnCSVFileSelected_WithEmptyPath_DoesNothing()
        {
            // Act
            _controller.OnCSVFileSelected("");

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(It.IsAny<string>()), Times.Never);
            _mockGui.Verify(g => g.DisplaySelectedCSVPath(It.IsAny<string>()), Times.Never);
        }

        #endregion

        #region OnExcelFileSelected Tests

        [Test]
        public void OnExcelFileSelected_WithValidFile_StoresPathAndEnablesProcessButton()
        {
            // Arrange
            string excelPath = "test.xlsx";
            var mockWorkbook = new Mock<ExcelWorkbook>(null);
            var mockFissiSheet = new Mock<Sheet>(null);

            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Returns(mockWorkbook.Object);
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook.Object))
                .Returns(mockFissiSheet.Object);

            // Act
            _controller.OnExcelFileSelected(excelPath);

            // Assert
            _mockGui.Verify(g => g.DisplaySelectedExcelPath(excelPath), Times.Once);
            Assert.That(_controller.CanProcess, Is.False); // CSV not selected yet
        }

        [Test]
        public void OnExcelFileSelected_WithNonExistentFile_ShowsError()
        {
            // Arrange
            string excelPath = "nonexistent.xlsx";
            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Throws<FileNotFoundException>();

            // Act
            _controller.OnExcelFileSelected(excelPath);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorExcelFileRead), Times.Once);
            _mockGui.Verify(g => g.DisplaySelectedExcelPath(It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnExcelFileSelected_WithoutFissiSheet_ShowsError()
        {
            // Arrange
            string excelPath = "test.xlsx";
            var mockWorkbook = new Mock<ExcelWorkbook>(null);

            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Returns(mockWorkbook.Object);
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook.Object))
                .Throws(new InvalidOperationException("Fissi sheet not found"));

            // Act
            _controller.OnExcelFileSelected(excelPath);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorFissiSheetNotFound), Times.Once);
            _mockGui.Verify(g => g.DisplaySelectedExcelPath(It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnExcelFileSelected_WithNullPath_DoesNothing()
        {
            // Act
            _controller.OnExcelFileSelected(null);

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(It.IsAny<string>()), Times.Never);
            _mockGui.Verify(g => g.DisplaySelectedExcelPath(It.IsAny<string>()), Times.Never);
        }

        #endregion

        #region CanProcess Tests

        [Test]
        public void CanProcess_WithBothFilesSelected_ReturnsTrue()
        {
            // Arrange
            string csvPath = Path.GetTempFileName();
            string excelPath = "test.xlsx";
            File.WriteAllText(csvPath, "test");
            
            var mockWorkbook = new Mock<ExcelWorkbook>(null);
            var mockFissiSheet = new Mock<Sheet>(null);

            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out It.Ref<List<string>>.IsAny))
                .Returns(true);
            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Returns(mockWorkbook.Object);
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook.Object))
                .Returns(mockFissiSheet.Object);

            // Act
            _controller.OnCSVFileSelected(csvPath);
            _controller.OnExcelFileSelected(excelPath);

            // Assert
            Assert.That(_controller.CanProcess, Is.True);

            // Cleanup
            File.Delete(csvPath);
        }

        [Test]
        public void CanProcess_WithOnlyCSVSelected_ReturnsFalse()
        {
            // Arrange
            string csvPath = Path.GetTempFileName();
            File.WriteAllText(csvPath, "test");
            
            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out It.Ref<List<string>>.IsAny))
                .Returns(true);

            // Act
            _controller.OnCSVFileSelected(csvPath);

            // Assert
            Assert.That(_controller.CanProcess, Is.False);

            // Cleanup
            File.Delete(csvPath);
        }

        [Test]
        public void CanProcess_WithOnlyExcelSelected_ReturnsFalse()
        {
            // Arrange
            string excelPath = "test.xlsx";
            var mockWorkbook = new Mock<ExcelWorkbook>(null);
            var mockFissiSheet = new Mock<Sheet>(null);

            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Returns(mockWorkbook.Object);
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook.Object))
                .Returns(mockFissiSheet.Object);

            // Act
            _controller.OnExcelFileSelected(excelPath);

            // Assert
            Assert.That(_controller.CanProcess, Is.False);
        }

        #endregion

        #region OnProcessButtonClicked Tests

        [Test]
        public void OnProcessButtonClicked_WithValidData_CreatesNewSheetSuccessfully()
        {
            // Arrange
            SetupValidFilesAndWorkflow();

            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockCsvParser.Verify(p => p.ParseCSV(It.IsAny<string>()), Times.Once);
            _mockDataTransformer.Verify(t => t.Transform(It.IsAny<List<ServiceAppointment>>()), Times.Once);
            _mockExcelManager.Verify(m => m.CreateNewSheet(It.IsAny<ExcelWorkbook>(), It.IsAny<int>()), Times.Once);
            _mockExcelManager.Verify(m => m.WriteHeader(It.IsAny<Sheet>(), It.IsAny<string>()), Times.Once);
            _mockExcelManager.Verify(m => m.WriteDataRows(It.IsAny<Sheet>(), It.IsAny<List<TransformedRow>>(), 2), Times.Once);
            _mockExcelManager.Verify(m => m.AppendFissiData(It.IsAny<Sheet>(), It.IsAny<Sheet>(), It.IsAny<int>(, DateTime.Now)), Times.Once);
            _mockExcelManager.Verify(m => m.ApplyYellowHighlight(It.IsAny<Sheet>(), It.IsAny<List<int>>()), Times.Once);
            _mockGui.Verify(g => g.EnableDownloadButton(true), Times.Once);
            Assert.That(_controller.CanDownload, Is.True);
        }

        [Test]
        public void OnProcessButtonClicked_WithCSVParseError_ShowsError()
        {
            // Arrange
            SetupValidFiles();
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Throws<FileNotFoundException>();

            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorCSVFileRead), Times.Once);
            _mockGui.Verify(g => g.EnableDownloadButton(true), Times.Never);
            Assert.That(_controller.CanDownload, Is.False);
        }

        [Test]
        public void OnProcessButtonClicked_WithMalformedCSV_ShowsError()
        {
            // Arrange
            SetupValidFiles();
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Throws<Exception>();

            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorCSVMalformed), Times.Once);
            Assert.That(_controller.CanDownload, Is.False);
        }

        [Test]
        public void OnProcessButtonClicked_WithTransformationError_ShowsError()
        {
            // Arrange
            SetupValidFiles();
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Returns(new List<ServiceAppointment>());
            _mockDataTransformer.Setup(t => t.Transform(It.IsAny<List<ServiceAppointment>>()))
                .Throws<Exception>();

            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorDataTransformation), Times.Once);
            Assert.That(_controller.CanDownload, Is.False);
        }

        [Test]
        public void OnProcessButtonClicked_WithHeaderParsingError_ShowsError()
        {
            // Arrange
            SetupValidFiles();
            SetupCSVAndTransformation();
            SetupExcelSheets();
            
            _mockHeaderCalculator.Setup(h => h.GenerateNextWeekHeader(It.IsAny<string>()))
                .Throws<FormatException>();

            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorHeaderParsing), Times.Once);
            Assert.That(_controller.CanDownload, Is.False);
        }

        [Test]
        public void OnProcessButtonClicked_WithoutFilesSelected_DoesNothing()
        {
            // Act
            _controller.OnProcessButtonClicked();

            // Assert
            _mockCsvParser.Verify(p => p.ParseCSV(It.IsAny<string>()), Times.Never);
            Assert.That(_controller.CanDownload, Is.False);
        }

        [Test]
        public void OnProcessButtonClicked_ApplicationRemainsUsableAfterError()
        {
            // Arrange
            SetupValidFiles();
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Throws<Exception>();

            // Act - First attempt fails
            _controller.OnProcessButtonClicked();

            // Assert - Application should still be in usable state
            Assert.That(_controller.CanProcess, Is.True);
            Assert.That(_controller.CanDownload, Is.False);

            // Act - Can try again with valid data
            SetupValidFilesAndWorkflow();
            _controller.OnProcessButtonClicked();

            // Assert - Should succeed
            Assert.That(_controller.CanDownload, Is.True);
        }

        #endregion

        #region OnDownloadButtonClicked Tests

        [Test]
        public void OnDownloadButtonClicked_WithProcessedWorkbook_SavesFile()
        {
            // Arrange
            SetupValidFilesAndWorkflow();
            _controller.OnProcessButtonClicked();
            
            string savePath = "output.xlsx";
            _mockGui.Setup(g => g.GetSaveFilePath()).Returns(savePath);

            // Act
            _controller.OnDownloadButtonClicked();

            // Assert
            _mockExcelManager.Verify(m => m.SaveWorkbook(It.IsAny<ExcelWorkbook>(), savePath), Times.Once);
            _mockGui.Verify(g => g.ShowSuccessMessage(Resources.SuccessMessage), Times.Once);
        }

        [Test]
        public void OnDownloadButtonClicked_WithSaveError_ShowsError()
        {
            // Arrange
            SetupValidFilesAndWorkflow();
            _controller.OnProcessButtonClicked();
            
            string savePath = "output.xlsx";
            _mockGui.Setup(g => g.GetSaveFilePath()).Returns(savePath);
            _mockExcelManager.Setup(m => m.SaveWorkbook(It.IsAny<ExcelWorkbook>(), savePath))
                .Throws<Exception>();

            // Act
            _controller.OnDownloadButtonClicked();

            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.ErrorFileSave), Times.Once);
            _mockGui.Verify(g => g.ShowSuccessMessage(It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnDownloadButtonClicked_WithUserCancellation_DoesNotSave()
        {
            // Arrange
            SetupValidFilesAndWorkflow();
            _controller.OnProcessButtonClicked();
            
            _mockGui.Setup(g => g.GetSaveFilePath()).Returns((string)null);

            // Act
            _controller.OnDownloadButtonClicked();

            // Assert
            _mockExcelManager.Verify(m => m.SaveWorkbook(It.IsAny<ExcelWorkbook>(), It.IsAny<string>()), Times.Never);
        }

        [Test]
        public void OnDownloadButtonClicked_WithoutProcessing_DoesNothing()
        {
            // Act
            _controller.OnDownloadButtonClicked();

            // Assert
            _mockGui.Verify(g => g.GetSaveFilePath(), Times.Never);
            _mockExcelManager.Verify(m => m.SaveWorkbook(It.IsAny<ExcelWorkbook>(), It.IsAny<string>()), Times.Never);
        }

        #endregion

        #region Helper Methods

        private void SetupValidFiles()
        {
            string csvPath = Path.GetTempFileName();
            string excelPath = "test.xlsx";
            File.WriteAllText(csvPath, "test");
            
            var mockWorkbook = new Mock<ExcelWorkbook>(null);
            var mockFissiSheet = new Mock<Sheet>(null);

            _mockCsvParser.Setup(p => p.ValidateCSVStructure(csvPath, out It.Ref<List<string>>.IsAny))
                .Returns(true);
            _mockExcelManager.Setup(m => m.OpenWorkbook(excelPath))
                .Returns(mockWorkbook.Object);
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook.Object))
                .Returns(mockFissiSheet.Object);

            _controller.OnCSVFileSelected(csvPath);
            _controller.OnExcelFileSelected(excelPath);

            File.Delete(csvPath);
        }

        private void SetupCSVAndTransformation()
        {
            var appointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "10:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario"
                }
            };

            var transformedRows = new List<TransformedRow>
            {
                new TransformedRow
                {
                    DataServizio = "2024-01-15",
                    OraInizioServizio = "10:00",
                    Assistito = "Rossi Mario"
                }
            };

            var transformationResult = new TransformationResult
            {
                Rows = transformedRows,
                YellowHighlightRows = new List<int> { 1 }
            };

            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Returns(appointments);
            _mockDataTransformer.Setup(t => t.Transform(appointments))
                .Returns(transformationResult);
        }

        private void SetupExcelSheets()
        {
            var mockWorkbook = new Mock<ExcelWorkbook>(null);
            var mockFissiSheet = new Mock<Sheet>(null);
            var mockLastSheet = new Mock<Sheet>(null);
            var mockNewSheet = new Mock<Sheet>(null);

            var sheetNames = new List<string> { "1", "2", "3", "fissi" };

            _mockExcelManager.Setup(m => m.GetSheetNames(It.IsAny<ExcelWorkbook>()))
                .Returns(sheetNames);
            _mockExcelManager.Setup(m => m.GetNextSheetNumber(sheetNames))
                .Returns(4);
            _mockExcelManager.Setup(m => m.GetFissiSheet(It.IsAny<ExcelWorkbook>()))
                .Returns(mockFissiSheet.Object);
            _mockExcelManager.Setup(m => m.GetSheetByName(It.IsAny<ExcelWorkbook>(), "3"))
                .Returns(mockLastSheet.Object);
            _mockExcelManager.Setup(m => m.ReadHeader(mockLastSheet.Object))
                .Returns("26 gen 01 feb Settimana 3referente settimana = Test");
            _mockExcelManager.Setup(m => m.CreateNewSheet(It.IsAny<ExcelWorkbook>(), 4))
                .Returns(mockNewSheet.Object);

            _mockHeaderCalculator.Setup(h => h.GenerateNextWeekHeader(It.IsAny<string>()))
                .Returns("02 feb 08 feb Settimana 4referente settimana = Inserire nome e numero di telefono del referente");
        }

        private void SetupValidFilesAndWorkflow()
        {
            SetupValidFiles();
            SetupCSVAndTransformation();
            SetupExcelSheets();
        }

        #endregion

        #region Enhanced Workflow Tests

        /// <summary>
        /// Test that OnProcessButtonClicked uses enhanced transformation with reference sheets
        /// Validates: Requirements 5.1, 5.2, 6.1, 6.2, 8.2, 8.3
        /// </summary>
        [Test]
        public void OnProcessButtonClicked_WithReferenceSheets_UsesEnhancedTransformation()
        {
            // Arrange
            SetupValidFiles();
            
            var mockAppointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2026-01-26",
                    OraInizioServizio = "09:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario"
                }
            };
            
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Returns(mockAppointments);
            
            var mockWorkbook = new ExcelWorkbook();
            _mockExcelManager.Setup(m => m.OpenWorkbook(It.IsAny<string>()))
                .Returns(mockWorkbook);
            
            var mockAssistitiSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetSheetByName(mockWorkbook, "assistiti"))
                .Returns(mockAssistitiSheet.Object);
            
            var mockFissiSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook))
                .Returns(mockFissiSheet.Object);
            
            _mockExcelManager.Setup(m => m.GetSheetNames(mockWorkbook))
                .Returns(new List<string> { "1", "2", "3" });
            _mockExcelManager.Setup(m => m.GetNextSheetNumber(It.IsAny<List<string>>()))
                .Returns(4);
            
            var mockLastSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetSheetByName(mockWorkbook, "3"))
                .Returns(mockLastSheet.Object);
            _mockExcelManager.Setup(m => m.ReadHeader(mockLastSheet.Object))
                .Returns("26 gen 01 feb Settimana 3referente settimana = Test");
            
            var mockHeaderInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2026, 1, 26),
                WeekNumber = 3
            };
            _mockHeaderCalculator.Setup(h => h.ParseHeader(It.IsAny<string>()))
                .Returns(mockHeaderInfo);
            
            var mockEnhancedResult = new EnhancedTransformationResult
            {
                Rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "26/01/2026",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario"
                    }
                },
                YellowHighlightRows = new List<int>()
            };
            
            _mockDataTransformer.Setup(t => t.TransformEnhanced(
                It.IsAny<List<ServiceAppointment>>(),
                It.IsAny<ILookupService>()))
                .Returns(mockEnhancedResult);
            
            var mockNewSheet = new Mock<Sheet>();
            var mockWorksheet = new Mock<OfficeOpenXml.ExcelWorksheet>();
            var mockDimension = new OfficeOpenXml.ExcelAddressBase(1, 1, 10, 15);
            mockWorksheet.Setup(w => w.Dimension).Returns(mockDimension);
            mockNewSheet.Setup(s => s.Worksheet).Returns(mockWorksheet.Object);
            
            _mockExcelManager.Setup(m => m.CreateNewSheet(mockWorkbook, 4))
                .Returns(mockNewSheet.Object);
            
            // Select files
            _controller.OnExcelFileSelected("test.xlsx");
            _controller.OnCSVFileSelected("test.csv");
            
            // Act
            _controller.OnProcessButtonClicked();
            
            // Assert
            _mockDataTransformer.Verify(t => t.TransformEnhanced(
                It.IsAny<List<ServiceAppointment>>(),
                It.IsAny<ILookupService>()), Times.Once);
            _mockExcelManager.Verify(m => m.WriteColumnHeadersEnhanced(It.IsAny<Sheet>()), Times.Once);
            _mockExcelManager.Verify(m => m.WriteDataRowsEnhanced(
                It.IsAny<Sheet>(),
                It.IsAny<List<EnhancedTransformedRow>>(),
                3), Times.Once);
        }

        /// <summary>
        /// Test that missing assistiti sheet shows appropriate error
        /// Validates: Requirements 5.1, 6.1
        /// </summary>
        [Test]
        public void OnExcelFileSelected_WithMissingAssistitiSheet_ShowsError()
        {
            // Arrange
            var mockWorkbook = new ExcelWorkbook();
            _mockExcelManager.Setup(m => m.OpenWorkbook(It.IsAny<string>()))
                .Returns(mockWorkbook);
            
            var mockFissiSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook))
                .Returns(mockFissiSheet.Object);
            
            _mockExcelManager.Setup(m => m.GetSheetByName(mockWorkbook, "assistiti"))
                .Returns((Sheet)null!);
            
            // Act
            _controller.OnExcelFileSelected("test.xlsx");
            
            // Assert
            _mockGui.Verify(g => g.ShowErrorMessage(Resources.AssistitiSheetNotFound), Times.Once);
            _mockGui.Verify(g => g.DisplaySelectedExcelPath(It.IsAny<string>()), Times.Never);
        }

        /// <summary>
        /// Test that OnProcessButtonClicked applies enhanced formatting
        /// Validates: Requirements 4.1, 9.1, 9.2, 10.1, 10.2, 10.3
        /// </summary>
        [Test]
        public void OnProcessButtonClicked_AppliesEnhancedFormatting()
        {
            // Arrange
            SetupValidFiles();
            
            var mockAppointments = new List<ServiceAppointment>
            {
                new ServiceAppointment
                {
                    DataServizio = "2026-01-26",
                    OraInizioServizio = "09:00",
                    CognomeAssistito = "Rossi",
                    NomeAssistito = "Mario"
                }
            };
            
            _mockCsvParser.Setup(p => p.ParseCSV(It.IsAny<string>()))
                .Returns(mockAppointments);
            
            var mockWorkbook = new ExcelWorkbook();
            _mockExcelManager.Setup(m => m.OpenWorkbook(It.IsAny<string>()))
                .Returns(mockWorkbook);
            
            var mockAssistitiSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetSheetByName(mockWorkbook, "assistiti"))
                .Returns(mockAssistitiSheet.Object);
            
            var mockFissiSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetFissiSheet(mockWorkbook))
                .Returns(mockFissiSheet.Object);
            
            _mockExcelManager.Setup(m => m.GetSheetNames(mockWorkbook))
                .Returns(new List<string> { "1", "2", "3" });
            _mockExcelManager.Setup(m => m.GetNextSheetNumber(It.IsAny<List<string>>()))
                .Returns(4);
            
            var mockLastSheet = new Mock<Sheet>();
            _mockExcelManager.Setup(m => m.GetSheetByName(mockWorkbook, "3"))
                .Returns(mockLastSheet.Object);
            _mockExcelManager.Setup(m => m.ReadHeader(mockLastSheet.Object))
                .Returns("26 gen 01 feb Settimana 3referente settimana = Test");
            
            var mockHeaderInfo = new HeaderInfo
            {
                MondayDate = new DateTime(2026, 1, 26),
                WeekNumber = 3
            };
            _mockHeaderCalculator.Setup(h => h.ParseHeader(It.IsAny<string>()))
                .Returns(mockHeaderInfo);
            
            var mockEnhancedResult = new EnhancedTransformationResult
            {
                Rows = new List<EnhancedTransformedRow>
                {
                    new EnhancedTransformedRow
                    {
                        Data = "26/01/2026",
                        Partenza = "09:00",
                        Assistito = "Rossi Mario"
                    }
                },
                YellowHighlightRows = new List<int>()
            };
            
            _mockDataTransformer.Setup(t => t.TransformEnhanced(
                It.IsAny<List<ServiceAppointment>>(),
                It.IsAny<ILookupService>()))
                .Returns(mockEnhancedResult);
            
            var mockNewSheet = new Mock<Sheet>();
            var mockWorksheet = new Mock<OfficeOpenXml.ExcelWorksheet>();
            var mockDimension = new OfficeOpenXml.ExcelAddressBase(1, 1, 10, 15);
            mockWorksheet.Setup(w => w.Dimension).Returns(mockDimension);
            mockNewSheet.Setup(s => s.Worksheet).Returns(mockWorksheet.Object);
            
            _mockExcelManager.Setup(m => m.CreateNewSheet(mockWorkbook, 4))
                .Returns(mockNewSheet.Object);
            
            // Select files
            _controller.OnExcelFileSelected("test.xlsx");
            _controller.OnCSVFileSelected("test.csv");
            
            // Act
            _controller.OnProcessButtonClicked();
            
            // Assert
            _mockExcelManager.Verify(m => m.ApplyBoldToHeaders(It.IsAny<Sheet>(), 2), Times.Once);
            _mockExcelManager.Verify(m => m.SortDataRows(It.IsAny<Sheet>(), 3, It.IsAny<int>()), Times.Once);
            _mockExcelManager.Verify(m => m.ApplyThickBordersToDateGroups(It.IsAny<Sheet>(), 3, It.IsAny<int>()), Times.Once);
        }

        #endregion
    }
}
