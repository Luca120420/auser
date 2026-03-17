using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Handles Gmail SMTP authentication and email sending operations.
/// </summary>
public class EmailService : IEmailService
{
    private const string GmailSmtpHost = "smtp.gmail.com";
    private const int GmailSmtpPort = 587;

    /// <summary>
    /// Sends an email notification to a volunteer with their assigned rows.
    /// </summary>
    /// <param name="toEmail">Recipient email address</param>
    /// <param name="volunteerSurname">Volunteer surname for personalization</param>
    /// <param name="assignedRows">List of assigned row data</param>
    /// <param name="credentials">Gmail credentials</param>
    /// <returns>True if sent successfully, false otherwise</returns>
    public async Task<bool> SendVolunteerNotificationAsync(
        string toEmail,
        string volunteerSurname,
        List<Dictionary<string, string>> assignedRows,
        GmailCredentials credentials)
    {
        try
        {
            // Create SMTP client with Gmail settings
            using var smtpClient = new SmtpClient(GmailSmtpHost, GmailSmtpPort)
            {
                EnableSsl = true,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(credentials.Email, credentials.AppPassword)
            };

            // Create mail message
            using var mailMessage = new MailMessage
            {
                From = new MailAddress(credentials.Email),
                Subject = "Auser notifica trasporti",
                Body = FormatEmailBody(volunteerSurname, assignedRows),
                IsBodyHtml = false
            };

            mailMessage.To.Add(toEmail);

            // Send email asynchronously
            await smtpClient.SendMailAsync(mailMessage);

            return true;
        }
        catch (SmtpException)
        {
            // Handle SMTP exceptions gracefully
            return false;
        }
        catch (Exception)
        {
            // Handle any other exceptions gracefully
            return false;
        }
    }

    /// <summary>
    /// Formats email body in Italian with assigned row data.
    /// Excludes columns: Volontario, Avv, Indirizzo Gasnet, Note Gasnet
    /// </summary>
    /// <param name="volunteerSurname">Volunteer surname</param>
    /// <param name="assignedRows">List of assigned row data</param>
    /// <returns>Formatted email body text</returns>
    public string FormatEmailBody(string volunteerSurname, List<Dictionary<string, string>> assignedRows)
    {
        var body = new StringBuilder();

        // Columns to exclude from email body
        var excludedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Volontario",
            "Avv",
            "Indirizzo Gasnet",
            "Note Gasnet"
        };

        // Italian greeting
        body.AppendLine($"Gentile {volunteerSurname},");
        body.AppendLine();
        body.AppendLine("Ecco i trasporti a te assegnati:");
        body.AppendLine();

        // Format each assigned row
        for (int i = 0; i < assignedRows.Count; i++)
        {
            var row = assignedRows[i];
            body.AppendLine($"Trasporto {i + 1}:");

            foreach (var column in row)
            {
                // Skip excluded columns and empty values
                if (!excludedColumns.Contains(column.Key) && !string.IsNullOrWhiteSpace(column.Value))
                {
                    body.AppendLine($"  {column.Key}: {column.Value}");
                }
            }

            // Add separator between rows (except after the last row)
            if (i < assignedRows.Count - 1)
            {
                body.AppendLine();
            }
        }

        body.AppendLine();
        body.AppendLine("Grazie per la tua disponibilità.");

        return body.ToString();
    }

    /// <summary>
    /// Tests Gmail SMTP connection with provided credentials.
    /// </summary>
    /// <param name="credentials">Gmail credentials to test</param>
    /// <returns>True if connection successful, false otherwise</returns>
    public async Task<bool> TestConnectionAsync(GmailCredentials credentials)
    {
        try
        {
            using var smtpClient = new SmtpClient(GmailSmtpHost, GmailSmtpPort)
            {
                EnableSsl = true,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(credentials.Email, credentials.AppPassword)
            };

            // Create a test message (not sent)
            using var testMessage = new MailMessage
            {
                From = new MailAddress(credentials.Email),
                Subject = "Test",
                Body = "Test"
            };
            testMessage.To.Add(credentials.Email);

            // Attempt to send to test connection
            await smtpClient.SendMailAsync(testMessage);

            return true;
        }
        catch (SmtpException)
        {
            return false;
        }
        catch (Exception)
        {
            return false;
        }
    }
}
