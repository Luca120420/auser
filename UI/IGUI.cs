namespace AuserExcelTransformer.UI
{
    /// <summary>
    /// Interface for the GUI layer.
    /// Defines methods for displaying information and getting user input.
    /// Validates: Requirements 1.1, 1.2, 1.3, 1.4, 1.5, 7.3, 7.4, 7.6, 8.1, 8.2, 8.3, 8.4, 8.5
    /// </summary>
    public interface IGUI
    {
        /// <summary>
        /// Shows the main application window.
        /// </summary>
        void ShowWindow();

        /// <summary>
        /// Opens a file selection dialog for CSV files.
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        string? SelectCSVFile();

        /// <summary>
        /// Opens a file selection dialog for Excel files.
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        string? SelectExcelFile();

        /// <summary>
        /// Displays the selected CSV file path in the GUI.
        /// </summary>
        /// <param name="path">The file path to display</param>
        void DisplaySelectedCSVPath(string path);

        /// <summary>
        /// Displays the selected Excel file path in the GUI.
        /// </summary>
        /// <param name="path">The file path to display</param>
        void DisplaySelectedExcelPath(string path);

        /// <summary>
        /// Enables or disables the process button.
        /// </summary>
        /// <param name="enabled">True to enable, false to disable</param>
        void EnableProcessButton(bool enabled);

        /// <summary>
        /// Enables or disables the download button.
        /// </summary>
        /// <param name="enabled">True to enable, false to disable</param>
        void EnableDownloadButton(bool enabled);

        /// <summary>
        /// Displays an error message to the user in Italian.
        /// </summary>
        /// <param name="message">The error message to display</param>
        void ShowErrorMessage(string message);

        /// <summary>
        /// Displays a success message to the user in Italian.
        /// </summary>
        /// <param name="message">The success message to display</param>
        void ShowSuccessMessage(string message);

        /// <summary>
        /// Opens a save file dialog and returns the selected path.
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        string? GetSaveFilePath();
    }
}
