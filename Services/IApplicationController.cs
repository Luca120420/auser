namespace AuserExcelTransformer.Services
{
    /// <summary>
    /// Interface for the application controller that orchestrates the workflow.
    /// Coordinates between CSV parsing, data transformation, Excel management, and header calculation.
    /// Validates: Requirements 7.1, 7.3, 7.4, 7.5, 7.6, 9.1, 9.2, 9.3, 9.4, 9.5, 9.6
    /// </summary>
    public interface IApplicationController
    {
        /// <summary>
        /// Handles the event when a CSV file is selected.
        /// Validates the CSV file structure and stores the path.
        /// </summary>
        /// <param name="csvPath">The path to the selected CSV file</param>
        void OnCSVFileSelected(string csvPath);

        /// <summary>
        /// Handles the event when an Excel file is selected.
        /// Opens the workbook and validates it has the required structure.
        /// </summary>
        /// <param name="excelPath">The path to the selected Excel file</param>
        void OnExcelFileSelected(string excelPath);

        /// <summary>
        /// Handles the event when the process button is clicked.
        /// Orchestrates the entire workflow:
        /// 1. Parse CSV data
        /// 2. Transform data
        /// 3. Get next sheet number
        /// 4. Get fissi sheet
        /// 5. Get last numbered sheet header
        /// 6. Generate next week header
        /// 7. Create new sheet
        /// 8. Write header and data
        /// 9. Append fissi data
        /// 10. Apply yellow highlighting
        /// </summary>
        void OnProcessButtonClicked();

        /// <summary>
        /// Handles the event when the download button is clicked.
        /// Prompts the user for a save location and saves the workbook.
        /// </summary>
        void OnDownloadButtonClicked();

        /// <summary>
        /// Gets whether both CSV and Excel files have been selected and validated.
        /// </summary>
        bool CanProcess { get; }

        /// <summary>
        /// Gets whether the workbook has been processed and is ready to download.
        /// </summary>
        bool CanDownload { get; }
    }
}
