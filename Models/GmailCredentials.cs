namespace AuserExcelTransformer.Models;

/// <summary>
/// Stores Gmail SMTP authentication credentials.
/// </summary>
public class GmailCredentials
{
    /// <summary>
    /// Gmail email address.
    /// </summary>
    public string Email { get; set; } = string.Empty;
    
    /// <summary>
    /// Gmail application password (not regular password).
    /// </summary>
    public string AppPassword { get; set; } = string.Empty;
    
    /// <summary>
    /// Checks if credentials are configured.
    /// </summary>
    public bool IsConfigured => !string.IsNullOrWhiteSpace(Email) && 
                                 !string.IsNullOrWhiteSpace(AppPassword);
}
