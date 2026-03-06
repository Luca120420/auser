namespace AuserExcelTransformer.Models;

/// <summary>
/// Stores application configuration data for persistence.
/// </summary>
public class AppConfiguration
{
    /// <summary>
    /// Gmail credentials for email sending.
    /// </summary>
    public GmailCredentials GmailCredentials { get; set; } = new GmailCredentials();
    
    /// <summary>
    /// Last selected Excel file path for notifications.
    /// </summary>
    public string LastExcelFilePath { get; set; } = string.Empty;
    
    /// <summary>
    /// Last selected sheet name.
    /// </summary>
    public string LastSheetName { get; set; } = string.Empty;
}
