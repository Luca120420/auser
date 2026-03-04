using System.Collections.Generic;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Interface for volunteer notification UI operations.
/// </summary>
public interface IVolunteerUI
{
    /// <summary>
    /// Displays the list of volunteer contacts.
    /// </summary>
    /// <param name="volunteers">Dictionary of surname to email mappings</param>
    void DisplayVolunteerList(Dictionary<string, string> volunteers);
    
    /// <summary>
    /// Displays Gmail credentials in the UI.
    /// </summary>
    /// <param name="email">Gmail email address</param>
    /// <param name="password">Gmail app password</param>
    void DisplayGmailCredentials(string email, string password);
    
    /// <summary>
    /// Displays available sheet names from Excel file.
    /// </summary>
    /// <param name="sheetNames">List of sheet names</param>
    void DisplaySheetNames(List<string> sheetNames);
    
    /// <summary>
    /// Enables or disables the send emails button.
    /// </summary>
    /// <param name="enabled">True to enable, false to disable</param>
    void EnableSendEmailsButton(bool enabled);
    
    /// <summary>
    /// Shows email sending progress.
    /// </summary>
    /// <param name="message">Progress message</param>
    void ShowEmailProgress(string message);
    
    /// <summary>
    /// Shows email sending summary.
    /// </summary>
    /// <param name="successCount">Number of successful sends</param>
    /// <param name="failureCount">Number of failed sends</param>
    void ShowEmailSummary(int successCount, int failureCount);
    
    /// <summary>
    /// Prompts user for confirmation.
    /// </summary>
    /// <param name="message">Confirmation message</param>
    /// <returns>True if confirmed, false otherwise</returns>
    bool ConfirmAction(string message);
    
    /// <summary>
    /// Shows an error message to the user.
    /// </summary>
    /// <param name="message">Error message to display</param>
    void ShowErrorMessage(string message);
}
