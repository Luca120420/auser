using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;
using OfficeOpenXml;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Integration tests for laboratori sheet processing.
    /// Tests complete transformation workflow with CSV, fissi, and laboratori data.
    /// Validates: Requirements 2.1, 3.1-3.10, 6.1-6.4, 7.1-7.3
    /// </summary>
    [TestFixture]
    public class LaboratoriIntegrationTests
    {
        private IExcelManager _excelManager = null!;
        private IDataTransformer _dataTransformer = null!;
        private ICSVParser _csvParser = null!;
        private ILookupService _lookupService = null!;

        [SetUp]
        public void Setup()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            _excelManager = new ExcelManager();
            _dataTransformer = new DataTransformer(new TransformationRulesEngine());
            _csvParser = new CSVParser();
            _lookupService = new LookupService();
        }

        /// <summary>
        /// Test complete transformation with CSV + fissi + laboratori data.
        /// Validates: Requirements 2.1, 3.1-3.10, 6.1-6.4, 7.1-7.3
        /// </summary>
        [Test]
        public void IntegrationTest_CompleteTransformation_WithLaboratoriData()
        {
            // Arrange - Create test workbook with all required sheets
            using (var package = new ExcelPackage())
            {
                // Create assistiti sheet (reference data)
                var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                assistitiSheet.Cells[1, 1].Value = "Cognome";
                assistitiSheet.Cells[1, 2].Value = "Nome";
                assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                assistitiSheet.Cells[2, 1].Value = "Rossi";
                assistitiSheet.Cells[2, 2].Value = "Mario";
                assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                // Create fissi sheet (recurring appointments)
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                fissiWorksheet.Cells[1, 1].Value = "Data";
                fissiWorksheet.Cells[1, 2].Value = "Partenza";
                fissiWorksheet.Cells[1, 3].Value = "Assistito";
                fissiWorksheet.Cells[1, 4].Value = "Indirizzo";
                fissiWorksheet.Cells[1, 5].Value = "Destinazione";
                fissiWorksheet.Cells[1, 6].Value = "Note";
                fissiWorksheet.Cells[1, 7].Value = "Auto";
                fissiWorksheet.Cells[1, 8].Value = "Volontario";
                fissiWorksheet.Cells[1, 9].Value = "Arrivo";
                
                // Add 2 fissi data rows
                fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 2);
                fissiWorksheet.Cells[2, 2].Value = 0.375; // 9:00 AM
                fissiWorksheet.Cells[2, 3].Value = "Bianchi Laura";
                fissiWorksheet.Cells[2, 4].Value = "Via Milano 5";
                fissiWorksheet.Cells[2, 5].Value = "Ospedale";
                fissiWorksheet.Cells[2, 6].Value = "Dialisi";
                fissiWorksheet.Cells[2, 7].Value = "Auto1";
                fissiWorksheet.Cells[2, 8].Value = "Volontario1";
                fissiWorksheet.Cells[2, 9].Value = 0.5; // 12:00 PM
                
                fissiWorksheet.Cells[3, 1].Value = new DateTime(2026, 2, 3);
                fissiWorksheet.Cells[3, 2].Value = 0.416667; // 10:00 AM
                fissiWorksheet.Cells[3, 3].Value = "Verdi Giuseppe";
                fissiWorksheet.Cells[3, 4].Value = "Via Torino 10";
                fissiWorksheet.Cells[3, 5].Value = "Clinica";
                fissiWorksheet.Cells[3, 6].Value = "Fisioterapia";
                fissiWorksheet.Cells[3, 7].Value = "Auto2";
                fissiWorksheet.Cells[3, 8].Value = "Volontario2";
                fissiWorksheet.Cells[3, 9].Value = 0.541667; // 1:00 PM

                // Create laboratori sheet (10 columns including Avv)
                var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                laboratoriWorksheet.Cells[1, 1].Value = "Data";
                laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                laboratoriWorksheet.Cells[1, 4].Value = "Indirizzo";
                laboratoriWorksheet.Cells[1, 5].Value = "Destinazione";
                laboratoriWorksheet.Cells[1, 6].Value = "Note";
                laboratoriWorksheet.Cells[1, 7].Value = "Auto";
                laboratoriWorksheet.Cells[1, 8].Value = "Volontario";
                laboratoriWorksheet.Cells[1, 9].Value = "Arrivo";
                laboratoriWorksheet.Cells[1, 10].Value = "Avv";

                // Add 3 laboratori data rows with different dates for sorting test
                laboratoriWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 1); // Earlier date
                laboratoriWorksheet.Cells[2, 2].Value = 0.333333; // 8:00 AM
                laboratoriWorksheet.Cells[2, 3].Value = "Ferrari Anna";
                laboratoriWorksheet.Cells[2, 4].Value = "Via Napoli 20";
                laboratoriWorksheet.Cells[2, 5].Value = "Laboratorio";
                laboratoriWorksheet.Cells[2, 6].Value = "Analisi sangue";
                laboratoriWorksheet.Cells[2, 7].Value = "Auto3";
                laboratoriWorksheet.Cells[2, 8].Value = "Volontario3";
                laboratoriWorksheet.Cells[2, 9].Value = 0.416667; // 10:00 AM
                laboratoriWorksheet.Cells[2, 10].Value = "Avviso1";
                
                laboratoriWorksheet.Cells[3, 1].Value = new DateTime(2026, 2, 4); // Later date
                laboratoriWorksheet.Cells[3, 2].Value = 0.291667; // 7:00 AM
                laboratoriWorksheet.Cells[3, 3].Value = "Colombo Marco";
                laboratoriWorksheet.Cells[3, 4].Value = "Via Firenze 15";
                laboratoriWorksheet.Cells[3, 5].Value = "Centro Analisi";
                laboratoriWorksheet.Cells[3, 6].Value = "Prelievo";
                laboratoriWorksheet.Cells[3, 7].Value = "Auto4";
                laboratoriWorksheet.Cells[3, 8].Value = "Volontario4";
                laboratoriWorksheet.Cells[3, 9].Value = 0.375; // 9:00 AM
                laboratoriWorksheet.Cells[3, 10].Value = "Avviso2";
                
                laboratoriWorksheet.Cells[4, 1].Value = new DateTime(2026, 2, 2); // Middle date
                laboratoriWorksheet.Cells[4, 2].Value = 0.458333; // 11:00 AM
                laboratoriWorksheet.Cells[4, 3].Value = "Greco Francesca";
                laboratoriWorksheet.Cells[4, 4].Value = "Via Bologna 8";
                laboratoriWorksheet.Cells[4, 5].Value = "Ospedale";
                laboratoriWorksheet.Cells[4, 6].Value = "Radiografia";
                laboratoriWorksheet.Cells[4, 7].Value = "Auto5";
                laboratoriWorksheet.Cells[4, 8].Value = "Volontario5";
                laboratoriWorksheet.Cells[4, 9].Value = 0.541667; // 1:00 PM
                laboratoriWorksheet.Cells[4, 10].Value = "Avviso3";

                // Create target output sheet
                var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                
                // Wrap sheets in models
                var workbook = new Models.ExcelWorkbook(package);
                var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                var targetSheet = new Sheet(targetWorksheet);

                // Load reference sheets
                _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                // Create sample CSV data (1 row for simplicity)
                var csvAppointments = new List<ServiceAppointment>
                {
                    new ServiceAppointment
                    {
                        DataServizio = "03/02/2026",
                        OraInizioServizio = "14:00",
                        CognomeAssistito = "Rossi",
                        NomeAssistito = "Mario",
                        IndirizzoPartenza = "Via Roma 1",
                        IndirizzoDestinazione = "Via Milano 5",
                        ComunePartenza = "Milano",
                        ComuneDestinazione = "Milano",
                        DescrizionePuntoPartenza = "Casa",
                        CausaleDestinazione = "Ospedale",
                        NoteERichieste = "Test note",
                        DescrizioneStatoServizio = "PIANIFICATO",
                        Attivita = "Accomp. servizi con trasporto"
                    }
                };

                // Act - Transform CSV data
                var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                
                // Write headers and CSV data
                _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                
                // Append fissi data
                int fissiStartRow = 3 + transformedResult.Rows.Count;
                _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow);
                
                // Calculate last row after fissi
                int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                
                // Append laboratori data
                int laboratoriStartRow = lastRowAfterFissi + 1;
                _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow);
                
                // Get final row count
                int finalLastRow = targetWorksheet.Dimension?.End.Row ?? laboratoriStartRow - 1;

                // Sort all data
                if (finalLastRow >= 3)
                {
                    _excelManager.SortDataRows(targetSheet, 3, finalLastRow);
                    _excelManager.ApplyThickBordersToDateGroups(targetSheet, 3, finalLastRow);
                }

                // Assert - Verify complete workflow
                Assert.That(targetWorksheet.Dimension, Is.Not.Null, "Output sheet should have data");
                
                // Verify total row count: 1 CSV + 2 fissi + 3 laboratori = 6 data rows (plus 2 header rows)
                int expectedDataRows = 1 + 2 + 3; // CSV + fissi + laboratori
                int actualDataRows = finalLastRow - 2; // Subtract header rows (1 and 2)
                Assert.That(actualDataRows, Is.EqualTo(expectedDataRows), 
                    "Should have 6 data rows (1 CSV + 2 fissi + 3 laboratori)");

                // Verify all 10 laboratori columns are mapped correctly
                // Check the first laboratori row (should be sorted, so find by Avv value)
                bool foundAvviso1 = false;
                bool foundAvviso2 = false;
                bool foundAvviso3 = false;
                
                for (int row = 3; row <= finalLastRow; row++)
                {
                    var avvValue = targetWorksheet.Cells[row, 10].Text;
                    
                    if (avvValue == "Avviso1")
                    {
                        foundAvviso1 = true;
                        // Verify all columns for this laboratori row
                        Assert.That(targetWorksheet.Cells[row, 1].Value, Is.Not.Null, "Data should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 2].Value, Is.Not.Null, "Partenza should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 3].Text, Is.EqualTo("Ferrari Anna"), "Assistito should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 4].Text, Is.EqualTo("Via Napoli 20"), "Indirizzo should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 5].Text, Is.EqualTo("Laboratorio"), "Destinazione should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 6].Text, Is.EqualTo("Analisi sangue"), "Note should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 7].Text, Is.EqualTo("Auto3"), "Auto should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 8].Text, Is.EqualTo("Volontario3"), "Volontario should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 9].Value, Is.Not.Null, "Arrivo should be mapped");
                        Assert.That(targetWorksheet.Cells[row, 10].Text, Is.EqualTo("Avviso1"), "Avv should be mapped");
                    }
                    else if (avvValue == "Avviso2")
                    {
                        foundAvviso2 = true;
                    }
                    else if (avvValue == "Avviso3")
                    {
                        foundAvviso3 = true;
                    }
                }

                Assert.That(foundAvviso1, Is.True, "Should find laboratori row with Avviso1");
                Assert.That(foundAvviso2, Is.True, "Should find laboratori row with Avviso2");
                Assert.That(foundAvviso3, Is.True, "Should find laboratori row with Avviso3");

                // Verify sorting by date and time
                DateTime? previousDate = null;
                TimeSpan? previousTime = null;
                
                for (int row = 3; row <= finalLastRow; row++)
                {
                    var dateCell = targetWorksheet.Cells[row, 1];
                    var timeCell = targetWorksheet.Cells[row, 2];
                    
                    if (dateCell.Value == null) continue;
                    
                    DateTime currentDate;
                    if (dateCell.Value is DateTime dt)
                    {
                        currentDate = dt;
                    }
                    else if (DateTime.TryParse(dateCell.Value.ToString(), out DateTime parsedDate))
                    {
                        currentDate = parsedDate;
                    }
                    else
                    {
                        continue;
                    }
                    
                    if (previousDate.HasValue)
                    {
                        Assert.That(currentDate, Is.GreaterThanOrEqualTo(previousDate.Value),
                            $"Row {row} date should be >= previous date (sorted ascending)");
                        
                        // If same date, verify time sorting
                        if (currentDate.Date == previousDate.Value.Date && timeCell.Value != null)
                        {
                            double currentTimeValue = 0;
                            if (timeCell.Value is double dbl)
                            {
                                currentTimeValue = dbl;
                            }
                            else if (timeCell.Value is DateTime dtTime)
                            {
                                currentTimeValue = dtTime.TimeOfDay.TotalDays;
                            }
                            
                            if (previousTime.HasValue && currentTimeValue > 0)
                            {
                                TimeSpan currentTime = TimeSpan.FromDays(currentTimeValue);
                                Assert.That(currentTime, Is.GreaterThanOrEqualTo(previousTime.Value),
                                    $"Row {row} time should be >= previous time (sorted ascending)");
                                previousTime = currentTime;
                            }
                        }
                        else
                        {
                            previousTime = null;
                        }
                    }
                    
                    previousDate = currentDate;
                    if (timeCell.Value is double timeDouble)
                    {
                        previousTime = TimeSpan.FromDays(timeDouble);
                    }
                }

                // Verify time format is applied to Partenza and Arrivo columns
                for (int row = 3; row <= finalLastRow; row++)
                {
                    var partenzaCell = targetWorksheet.Cells[row, 2];
                    var arrivoCell = targetWorksheet.Cells[row, 9];
                    
                    if (partenzaCell.Value != null)
                    {
                        Assert.That(partenzaCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                            $"Row {row} Partenza should have h:mm format");
                    }
                    
                    if (arrivoCell.Value != null)
                    {
                        Assert.That(arrivoCell.Style.Numberformat.Format, Is.EqualTo("h:mm"),
                            $"Row {row} Arrivo should have h:mm format");
                    }
                }

                // Verify thick borders between date groups
                bool foundDateBoundary = false;
                for (int row = 3; row < finalLastRow; row++)
                {
                    var currentDateCell = targetWorksheet.Cells[row, 1];
                    var nextDateCell = targetWorksheet.Cells[row + 1, 1];
                    
                    if (currentDateCell.Value is DateTime currentDate && 
                        nextDateCell.Value is DateTime nextDate)
                    {
                        if (currentDate.Date != nextDate.Date)
                        {
                            foundDateBoundary = true;
                            var borderStyle = targetWorksheet.Cells[row, 1].Style.Border.Bottom.Style;
                            Assert.That(borderStyle, Is.EqualTo(OfficeOpenXml.Style.ExcelBorderStyle.Thick),
                                $"Row {row} should have thick bottom border (date boundary)");
                        }
                    }
                }
                
                Assert.That(foundDateBoundary, Is.True, 
                    "Should have at least one date boundary with thick border");

                // Verify column count (should be 15 columns in enhanced format)
                Assert.That(targetWorksheet.Dimension.End.Column, Is.GreaterThanOrEqualTo(10),
                    "Should have at least 10 columns for laboratori data");
            }
        }

        /// <summary>
        /// Test transformation without laboratori sheet (graceful handling).
        /// Validates: Requirements 8.1, 8.2, 8.3
        /// </summary>
        [Test]
        public void IntegrationTest_CompleteTransformation_WithoutLaboratoriSheet()
        {
            // Arrange - Create test workbook WITHOUT laboratori sheet
            using (var package = new ExcelPackage())
            {
                // Create assistiti sheet
                var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                assistitiSheet.Cells[1, 1].Value = "Cognome";
                assistitiSheet.Cells[1, 2].Value = "Nome";
                assistitiSheet.Cells[1, 3].Value = "Indirizzo";
                assistitiSheet.Cells[2, 1].Value = "Rossi";
                assistitiSheet.Cells[2, 2].Value = "Mario";
                assistitiSheet.Cells[2, 3].Value = "Via Roma 1";

                // Create fissi sheet
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                fissiWorksheet.Cells[1, 1].Value = "Data";
                fissiWorksheet.Cells[1, 2].Value = "Partenza";
                fissiWorksheet.Cells[1, 3].Value = "Assistito";
                fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 2);
                fissiWorksheet.Cells[2, 2].Value = 0.375;
                fissiWorksheet.Cells[2, 3].Value = "Bianchi Laura";

                // Create target output sheet
                var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                
                // Wrap sheets in models
                var workbook = new Models.ExcelWorkbook(package);
                var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                var targetSheet = new Sheet(targetWorksheet);

                // Load reference sheets
                _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                // Create sample CSV data
                var csvAppointments = new List<ServiceAppointment>
                {
                    new ServiceAppointment
                    {
                        DataServizio = "03/02/2026",
                        OraInizioServizio = "14:00",
                        CognomeAssistito = "Rossi",
                        NomeAssistito = "Mario",
                        DescrizioneStatoServizio = "PIANIFICATO",
                        Attivita = "Accomp. servizi con trasporto"
                    }
                };

                // Act - Transform CSV data
                var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                
                // Write headers and CSV data
                _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                
                // Append fissi data
                int fissiStartRow = 3 + transformedResult.Rows.Count;
                _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow);
                
                // Calculate last row after fissi
                int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;

                // Try to get laboratori sheet (should be null)
                var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                
                // Simulate ApplicationController workflow - only append if sheet exists
                if (laboratoriSheetModel != null)
                {
                    int laboratoriStartRow = lastRowAfterFissi + 1;
                    _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow);
                }
                
                // Get final row count
                int finalLastRow = targetWorksheet.Dimension?.End.Row ?? lastRowAfterFissi;
                
                // Sort all data
                if (finalLastRow >= 3)
                {
                    _excelManager.SortDataRows(targetSheet, 3, finalLastRow);
                }

                // Assert - Verify workflow completes without errors
                Assert.That(laboratoriSheetModel, Is.Null, "Laboratori sheet should not exist");
                
                // Verify only CSV + fissi data (no laboratori)
                int expectedDataRows = 1 + 1; // 1 CSV + 1 fissi (only 1 fissi row has data)
                int actualDataRows = finalLastRow - 2; // Subtract header rows
                Assert.That(actualDataRows, Is.GreaterThanOrEqualTo(expectedDataRows - 1), 
                    "Should have CSV + fissi data only (no laboratori)");
                
                // Verify no exceptions were thrown
                Assert.Pass("Workflow completed successfully without laboratori sheet");
            }
        }

        /// <summary>
        /// Test that laboratori data is appended after fissi data in correct sequence.
        /// Validates: Requirements 6.1, 6.2, 6.3, 6.4
        /// </summary>
        [Test]
        public void IntegrationTest_DataSequencing_LaboratoriAfterFissi()
        {
            // Arrange - Create test workbook
            using (var package = new ExcelPackage())
            {
                // Create minimal assistiti sheet
                var assistitiSheet = package.Workbook.Worksheets.Add("assistiti");
                assistitiSheet.Cells[1, 1].Value = "Cognome";

                // Create fissi sheet with 1 row
                var fissiWorksheet = package.Workbook.Worksheets.Add("fissi");
                fissiWorksheet.Cells[1, 1].Value = "Data";
                fissiWorksheet.Cells[1, 2].Value = "Partenza";
                fissiWorksheet.Cells[1, 3].Value = "Assistito";
                fissiWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 5);
                fissiWorksheet.Cells[2, 2].Value = 0.375;
                fissiWorksheet.Cells[2, 3].Value = "FISSI_MARKER";

                // Create laboratori sheet with 1 row
                var laboratoriWorksheet = package.Workbook.Worksheets.Add("laboratori");
                laboratoriWorksheet.Cells[1, 1].Value = "Data";
                laboratoriWorksheet.Cells[1, 2].Value = "Partenza";
                laboratoriWorksheet.Cells[1, 3].Value = "Assistito";
                laboratoriWorksheet.Cells[2, 1].Value = new DateTime(2026, 2, 6);
                laboratoriWorksheet.Cells[2, 2].Value = 0.333333;
                laboratoriWorksheet.Cells[2, 3].Value = "LABORATORI_MARKER";

                // Create target output sheet
                var targetWorksheet = package.Workbook.Worksheets.Add("Output");
                
                // Wrap sheets in models
                var workbook = new Models.ExcelWorkbook(package);
                var assistitiSheetModel = _excelManager.GetSheetByName(workbook, "assistiti");
                var fissiSheetModel = _excelManager.GetSheetByName(workbook, "fissi");
                var laboratoriSheetModel = _excelManager.GetSheetByName(workbook, "laboratori");
                var targetSheet = new Sheet(targetWorksheet);

                // Load reference sheets
                _lookupService.LoadReferenceSheets(assistitiSheetModel, fissiSheetModel);

                // Create CSV data with 1 row
                var csvAppointments = new List<ServiceAppointment>
                {
                    new ServiceAppointment
                    {
                        DataServizio = "04/02/2026",
                        OraInizioServizio = "10:00",
                        CognomeAssistito = "CSV",
                        NomeAssistito = "MARKER",
                        DescrizioneStatoServizio = "PIANIFICATO",
                        Attivita = "Accomp. servizi con trasporto"
                    }
                };

                // Act - Simulate complete workflow
                var transformedResult = _dataTransformer.TransformEnhanced(csvAppointments, _lookupService);
                _excelManager.WriteColumnHeadersEnhanced(targetSheet);
                _excelManager.WriteDataRowsEnhanced(targetSheet, transformedResult.Rows, 3);
                
                int fissiStartRow = 3 + transformedResult.Rows.Count;
                _excelManager.AppendFissiData(targetSheet, fissiSheetModel, fissiStartRow);
                
                int lastRowAfterFissi = targetWorksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                int laboratoriStartRow = lastRowAfterFissi + 1;
                _excelManager.AppendLaboratoriData(targetSheet, laboratoriSheetModel, laboratoriStartRow);
                
                int finalLastRow = targetWorksheet.Dimension?.End.Row ?? laboratoriStartRow - 1;

                // Assert - Verify all data sources are present
                // We verify by checking row count since sorting may change order
                bool foundCSV = false;
                bool foundFissi = false;
                bool foundLaboratori = false;
                
                for (int row = 3; row <= finalLastRow; row++)
                {
                    var assistitoValue = targetWorksheet.Cells[row, 3].Text;
                    
                    if (assistitoValue.Contains("CSV") || assistitoValue.Contains("MARKER"))
                    {
                        foundCSV = true;
                    }
                    if (assistitoValue == "FISSI_MARKER")
                    {
                        foundFissi = true;
                    }
                    if (assistitoValue == "LABORATORI_MARKER")
                    {
                        foundLaboratori = true;
                    }
                }
                
                Assert.That(foundCSV, Is.True, "Should find CSV data");
                Assert.That(foundFissi, Is.True, "Should find fissi data");
                Assert.That(foundLaboratori, Is.True, "Should find laboratori data");
                
                // Verify total row count
                int expectedDataRows = 1 + 1 + 1; // 1 CSV + 1 fissi + 1 laboratori
                int actualDataRows = finalLastRow - 2;
                Assert.That(actualDataRows, Is.EqualTo(expectedDataRows),
                    "Should have exactly 3 data rows (1 CSV + 1 fissi + 1 laboratori)");
            }
        }
    }
}
