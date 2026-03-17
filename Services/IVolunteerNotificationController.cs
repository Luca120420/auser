using System.Collections.Generic;
using System.Threading.Tasks;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Orchestrates the volunteer notification workflow.
/// Handles volunteer file selection, contact management, Gmail credentials,
/// Excel file/sheet selection, and email sending operations.
/// </summary>
public interface IVolunteerNotificationController
{
    /// <summary>
    /// Handles volunteer file selection and loading.
    /// </summary>
    /// <param name="filePath">Path to volunteer JSON file</param>
    void OnVolunteerFileSelected(string filePath);
    
    /// <summary>
    /// Handles adding a new volunteer contact.
    /// </summary>
    /// <param name="surname">Volunteer surname</param>
    /// <param name="email">Volunteer email address</param>
    void OnAddVolunteer(string surname, string email);
    
    /// <summary>
    /// Handles deleting a specific volunteer contact.
    /// </summary>
    /// <param name="surname">Volunteer surname to delete</param>
    void OnDeleteVolunteer(string surname);
    
    /// <summary>
    /// Handles deleting all volunteer contacts.
    /// </summary>
    void OnDeleteAllVolunteers();
    
    /// <summary>
    /// Handles Gmail credentials update (in-memory only, does not persist).
    /// </summary>
    /// <param name="email">Gmail email address</param>
    /// <param name="appPassword">Gmail application password</param>
    void OnGmailCredentialsUpdated(string email, string appPassword);

    /// <summary>
    /// Explicitly saves Gmail credentials to persistent storage.
    /// </summary>
    void SaveGmailCredentials();
    
    /// <summary>
    /// Handles clearing Gmail credentials.
    /// </summary>
    void OnClearGmailCredentials();
    
    /// <summary>
    /// Handles Excel file selection for volunteer notifications.
    /// </summary>
    /// <param name="filePath">Path to Excel file</param>
    void OnNotificationExcelFileSelected(string filePath);
    
    /// <summary>
    /// Handles sheet selection from Excel file.
    /// </summary>
    /// <param name="sheetName">Selected sheet name</param>
    void OnSheetSelected(string sheetName);
    
    /// <summary>
    /// Handles send emails button click.
    /// Identifies volunteer assignments and sends notifications.
    /// </summary>
    Task OnSendEmailsClickedAsync();
    
    /// <summary>
    /// Gets current volunteer contacts.
    /// </summary>
    /// <returns>Dictionary of surname to email mappings</returns>
    Dictionary<string, string> GetVolunteers();
    
    /// <summary>
    /// Gets current Gmail credentials.
    /// </summary>
    /// <returns>Tuple of email and app password</returns>
    (string Email, string AppPassword) GetGmailCredentials();
    
    /// <summary>
    /// Checks if email sending is enabled (all prerequisites met).
    /// </summary>
    /// <returns>True if can send emails, false otherwise</returns>
    bool CanSendEmails();
    
    /// <summary>
    /// Refreshes the UI display with current configuration data.
    /// Should be called after the UI is fully initialized.
    /// </summary>
    void RefreshUIDisplay();
}
