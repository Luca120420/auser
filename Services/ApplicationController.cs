using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Properties;
using AuserExcelTransformer.UI;

namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Application controller that orchestrates the workflow between GUI, CSV parsing,
    /// data transformation, Excel management, and header calculation.
    /// Validates: Requirements 7.1, 7.3, 7.4, 7.5, 7.6, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6
    /// </summary>
    public class ApplicationController : IApplicationController
    {
        private readonly IGUI _gui;
        private readonly ICSVParser _csvParser;
        private readonly IExcelManager _excelManager;
        private readonly IDataTransformer _dataTransformer;
        private readonly IHeaderCalculator _headerCalculator;

        private string? _csvPath;
        private string? _excelPath;
        private ExcelWorkbook? _workbook;
        private bool _isProcessed;

        /// <summary>
        /// Gets whether both CSV and Excel files have been selected and validated.
        /// </summary>
        public bool CanProcess => !string.IsNullOrEmpty(_csvPath) && !string.IsNullOrEmpty(_excelPath);

        /// <summary>
        /// Gets whether the workbook has been processed and is ready to download.
        /// </summary>
        public bool CanDownload => _isProcessed && _workbook != null;

        /// <summary>
        /// Initializes a new instance of the ApplicationController class.
        /// </summary>
        /// <param name="gui">The GUI interface</param>
        /// <param name="csvParser">The CSV parser service</param>
        /// <param name="excelManager">The Excel manager service</param>
        /// <param name="dataTransformer">The data transformer service</param>
        /// <param name="headerCalculator">The header calculator service</param>
        public ApplicationController(
            IGUI gui,
            ICSVParser csvParser,
            IExcelManager excelManager,
            IDataTransformer dataTransformer,
            IHeaderCalculator headerCalculator)
        {
            _gui = gui ?? throw new ArgumentNullException(nameof(gui));
            _csvParser = csvParser ?? throw new ArgumentNullException(nameof(csvParser));
            _excelManager = excelManager ?? throw new ArgumentNullException(nameof(excelManager));
            _dataTransformer = dataTransformer ?? throw new ArgumentNullException(nameof(dataTransformer));
            _headerCalculator = headerCalculator ?? throw new ArgumentNullException(nameof(headerCalculator));

            _isProcessed = false;
        }

        /// <summary>
        /// Handles the event when a CSV file is selected.
        /// Validates the CSV file structure and stores the path.
        /// Validates: Requirements 9.1, 9.3, 9.6
        /// </summary>
        /// <param name="csvPath">The path to the selected CSV file</param>
        public void OnCSVFileSelected(string csvPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(csvPath))
                {
                    return;
                }

                // Validate CSV structure
                if (!_csvParser.ValidateCSVStructure(csvPath, out List<string> missingColumns))
                {
                    string errorMessage = string.Format(Resources.ErrorCSVMissingColumns, string.Join(", ", missingColumns));
                    _gui.ShowErrorMessage(errorMessage);
                    return;
                }

                // Store the path and display it
                _csvPath = csvPath;
                _gui.DisplaySelectedCSVPath(csvPath);

                // Update process button state
                _gui.EnableProcessButton(CanProcess);
            }
            catch (FileNotFoundException)
            {
                _gui.ShowErrorMessage(Resources.ErrorCSVFileRead);
            }
            catch (IOException)
            {
                _gui.ShowErrorMessage(Resources.ErrorCSVFileRead);
            }
            catch (Exception ex)
            {
                _gui.ShowErrorMessage(string.Format(Resources.ErrorGeneral, ex.Message));
            }
        }

        /// <summary>
        /// Handles the event when an Excel file is selected.
        /// Opens the workbook and validates it has the required structure.
        /// Validates: Requirements 9.2, 9.5, 9.6
        /// </summary>
        /// <param name="excelPath">The path to the selected Excel file</param>
        public void OnExcelFileSelected(string excelPath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(excelPath))
                {
                    return;
                }

                // Try to open the workbook
                _workbook = _excelManager.OpenWorkbook(excelPath);

                // Validate that fissi sheet exists
                try
                {
                    _excelManager.GetFissiSheet(_workbook);
                }
                catch (InvalidOperationException)
                {
                    _gui.ShowErrorMessage(Resources.ErrorFissiSheetNotFound);
                    _workbook = null!;
                    return;
                }

                // Validate that assistiti sheet exists
                try
                {
                    var assistitiSheet = _excelManager.GetSheetByName(_workbook, "assistiti");
                    if (assistitiSheet == null)
                    {
                        _gui.ShowErrorMessage(Resources.AssistitiSheetNotFound);
                        _workbook = null!;
                        return;
                    }
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.AssistitiSheetNotFound);
                    _workbook = null!;
                    return;
                }

                // Store the path and display it
                _excelPath = excelPath;
                _gui.DisplaySelectedExcelPath(excelPath);

                // Update process button state
                _gui.EnableProcessButton(CanProcess);
            }
            catch (FileNotFoundException)
            {
                _gui.ShowErrorMessage(Resources.ErrorExcelFileRead);
            }
            catch (Exception ex)
            {
                _gui.ShowErrorMessage(string.Format(Resources.ErrorGeneral, ex.Message));
            }
        }

        /// <summary>
        /// Handles the event when the process button is clicked.
        /// Orchestrates the entire workflow.
        /// Validates: Requirements 7.1, 7.2, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6
        /// </summary>
        public void OnProcessButtonClicked()
        {
            try
            {
                if (!CanProcess)
                {
                    return;
                }

                // Step 1: Parse CSV data
                List<ServiceAppointment> appointments;
                try
                {
                    appointments = _csvParser.ParseCSV(_csvPath);
                }
                catch (FileNotFoundException)
                {
                    _gui.ShowErrorMessage(Resources.ErrorCSVFileRead);
                    return;
                }
                catch (IOException)
                {
                    _gui.ShowErrorMessage(Resources.ErrorCSVFileRead);
                    return;
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.ErrorCSVMalformed);
                    return;
                }

                // Step 2: Get sheet names and next sheet number
                List<string> sheetNames = _excelManager.GetSheetNames(_workbook);
                int nextSheetNumber = _excelManager.GetNextSheetNumber(sheetNames);

                // Step 3: Get reference sheets (assistiti and fissi)
                Sheet assistitiSheet;
                Sheet fissiSheet;
                try
                {
                    assistitiSheet = _excelManager.GetSheetByName(_workbook, "assistiti");
                    if (assistitiSheet == null)
                    {
                        _gui.ShowErrorMessage(Resources.AssistitiSheetNotFound);
                        return;
                    }

                    fissiSheet = _excelManager.GetFissiSheet(_workbook);
                }
                catch (InvalidOperationException)
                {
                    _gui.ShowErrorMessage(Resources.ErrorFissiSheetNotFound);
                    return;
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.AssistitiSheetNotFound);
                    return;
                }

                // Step 4: Initialize LookupService with reference sheets
                var lookupService = new LookupService();
                try
                {
                    lookupService.LoadReferenceSheets(assistitiSheet, fissiSheet);
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(string.Format(Resources.ReferenceSheetMalformed, "assistiti/fissi"));
                    return;
                }

                // Step 5: Transform data using enhanced transformation with lookups
                EnhancedTransformationResult enhancedTransformationResult;
                try
                {
                    enhancedTransformationResult = _dataTransformer.TransformEnhanced(appointments, lookupService);
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.ErrorDataTransformation);
                    return;
                }

                // Step 6: Get last numbered sheet and read its header
                string previousHeader;
                HeaderInfo previousHeaderInfo;
                try
                {
                    // Find the highest numbered sheet
                    int lastSheetNumber = nextSheetNumber - 1;
                    if (lastSheetNumber < 1)
                    {
                        _gui.ShowErrorMessage(Resources.ErrorHeaderParsing);
                        return;
                    }

                    Sheet lastSheet = _excelManager.GetSheetByName(_workbook, lastSheetNumber.ToString());
                    if (lastSheet == null)
                    {
                        _gui.ShowErrorMessage(Resources.ErrorHeaderParsing);
                        return;
                    }

                    previousHeader = _excelManager.ReadHeader(lastSheet);
                    previousHeaderInfo = _headerCalculator.ParseHeader(previousHeader);
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.ErrorHeaderParsing);
                    return;
                }

                // Step 6: Calculate next week's Monday date
                DateTime nextMondayDate;
                try
                {
                    // Calculate next Monday by adding 7 days to previous Monday
                    nextMondayDate = previousHeaderInfo.MondayDate.AddDays(7);
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.ErrorDateParsing);
                    return;
                }

                // Step 7: Create new sheet
                Sheet newSheet = _excelManager.CreateNewSheet(_workbook, nextSheetNumber);

                // Step 8: Write header row and column headers (enhanced)
                _excelManager.WriteHeaderRow(newSheet, nextMondayDate);
                _excelManager.WriteColumnHeadersEnhanced(newSheet);
                
                // Step 8.5: Apply bold formatting to headers
                _excelManager.ApplyBoldToHeaders(newSheet, 2);

                // Step 9: Write data rows using enhanced format (starting at row 3, after header and column headers)
                _excelManager.WriteDataRowsEnhanced(newSheet, enhancedTransformationResult.Rows, 3);

                // Step 10: Append fissi data
                int fissiStartRow = 3 + enhancedTransformationResult.Rows.Count;
                int fissiRowCountBefore = fissiStartRow;
                _excelManager.AppendFissiData(newSheet, fissiSheet, fissiStartRow);
                
                // Calculate the last data row by checking the worksheet dimension
                int lastDataRow = newSheet.Worksheet.Dimension?.End.Row ?? fissiStartRow - 1;
                
                // Step 10.5: Sort all data rows by date and time (if there are data rows)
                if (lastDataRow >= 3)
                {
                    _excelManager.SortDataRows(newSheet, 3, lastDataRow);
                    
                    // Step 10.6: Apply thick borders between date groups
                    _excelManager.ApplyThickBordersToDateGroups(newSheet, 3, lastDataRow);
                }

                // Step 11: Apply yellow highlighting to rows with "Accompag. con macchina attrezzata"
                _excelManager.ApplyYellowHighlight(newSheet, enhancedTransformationResult.YellowHighlightRows);

                // Step 12: Enable AutoFilter for sorting and filtering
                _excelManager.EnableAutoFilter(newSheet);

                // Mark as processed and enable download button
                _isProcessed = true;
                _gui.EnableDownloadButton(true);
                
                // Show success message
                _gui.ShowSuccessMessage(Resources.ProcessingComplete);
            }
            catch (Exception ex)
            {
                _gui.ShowErrorMessage(string.Format(Resources.ErrorGeneral, ex.Message));
            }
        }

        /// <summary>
        /// Handles the event when the download button is clicked.
        /// Prompts the user for a save location and saves the workbook.
        /// Validates: Requirements 7.4, 7.5, 7.6, 9.5, 9.6
        /// </summary>
        public void OnDownloadButtonClicked()
        {
            try
            {
                if (!CanDownload)
                {
                    return;
                }

                // Get save file path from user
                string savePath = _gui.GetSaveFilePath();
                if (string.IsNullOrWhiteSpace(savePath))
                {
                    return; // User cancelled
                }

                // Save the workbook
                try
                {
                    _excelManager.SaveWorkbook(_workbook, savePath);
                    _gui.ShowSuccessMessage(Resources.SuccessMessage);
                }
                catch (Exception)
                {
                    _gui.ShowErrorMessage(Resources.ErrorFileSave);
                }
            }
            catch (Exception ex)
            {
                _gui.ShowErrorMessage(string.Format(Resources.ErrorGeneral, ex.Message));
            }
        }
    }
}
