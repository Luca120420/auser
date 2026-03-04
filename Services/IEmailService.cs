using System.Collections.Generic;
using System.Threading.Tasks;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Handles Gmail SMTP authentication and email sending operations.
/// </summary>
public interface IEmailService
{
    /// <summary>
    /// Sends an email notification to a volunteer with their assigned rows.
    /// </summary>
    /// <param name="toEmail">Recipient email address</param>
    /// <param name="volunteerSurname">Volunteer surname for personalization</param>
    /// <param name="assignedRows">List of assigned row data</param>
    /// <param name="credentials">Gmail credentials</param>
    /// <returns>True if sent successfully, false otherwise</returns>
    Task<bool> SendVolunteerNotificationAsync(
        string toEmail, 
        string volunteerSurname, 
        List<Dictionary<string, string>> assignedRows,
        GmailCredentials credentials);
    
    /// <summary>
    /// Formats email body in Italian with assigned row data.
    /// </summary>
    /// <param name="volunteerSurname">Volunteer surname</param>
    /// <param name="assignedRows">List of assigned row data</param>
    /// <returns>Formatted email body text</returns>
    string FormatEmailBody(string volunteerSurname, List<Dictionary<string, string>> assignedRows);
    
    /// <summary>
    /// Tests Gmail SMTP connection with provided credentials.
    /// </summary>
    /// <param name="credentials">Gmail credentials to test</param>
    /// <returns>True if connection successful, false otherwise</returns>
    Task<bool> TestConnectionAsync(GmailCredentials credentials);
}
