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
                IsBodyHtml = true
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
    /// Formats email body as a modern, styled HTML email with assigned row data.
    /// Excludes columns: Volontario, Avv, Indirizzo Gasnet, Note Gasnet
    /// Uses inline CSS only (no external stylesheets/JS), since that's what email
    /// clients like Gmail/Outlook/Apple Mail reliably support. Text content (the
    /// greeting, section labels, and "Key: Value" pairs) is kept as unbroken text
    /// runs rather than being split across separate tags, so it stays readable.
    /// </summary>
    /// <param name="volunteerSurname">Volunteer surname</param>
    /// <param name="assignedRows">List of assigned row data</param>
    /// <returns>Formatted HTML email body</returns>
    public string FormatEmailBody(string volunteerSurname, List<Dictionary<string, string>> assignedRows)
    {
        // Columns to exclude from email body
        var excludedColumns = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "Volontario",
            "Avv",
            "Indirizzo Gasnet",
            "Note Gasnet"
        };

        const string primaryColor = "#2E7D32";   // brand green
        const string primaryDark = "#1B5E20";
        const string textColor = "#263238";
        const string mutedColor = "#607D8B";
        const string cardBorder = "#E0E0E0";
        const string pageBackground = "#F2F4F3";

        var body = new StringBuilder();

        body.Append("<!DOCTYPE html>");
        body.Append("<html lang=\"it\">");
        body.Append("<head><meta charset=\"UTF-8\"><meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\"></head>");
        body.Append($"<body style=\"margin:0; padding:0; background-color:{pageBackground}; font-family:'Segoe UI', Roboto, Helvetica, Arial, sans-serif;\">");

        // Outer wrapper table for consistent centering across email clients
        body.Append($"<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"background-color:{pageBackground}; padding:24px 0;\">");
        body.Append("<tr><td align=\"center\">");
        body.Append("<table role=\"presentation\" width=\"600\" cellpadding=\"0\" cellspacing=\"0\" style=\"max-width:600px; width:100%; background-color:#FFFFFF; border-radius:16px; overflow:hidden; box-shadow:0 2px 10px rgba(0,0,0,0.06);\">");

        // Header band
        body.Append("<tr><td style=\"background-color:" + primaryColor + "; padding:28px 32px;\">");
        body.Append("<div style=\"color:#FFFFFF; font-size:13px; letter-spacing:1.5px; text-transform:uppercase; opacity:0.85;\">Auser</div>");
        body.Append("<div style=\"color:#FFFFFF; font-size:22px; font-weight:700; margin-top:4px;\">Notifica Trasporti</div>");
        body.Append("</td></tr>");

        // Greeting — kept as a single unbroken text run ("Gentile {surname},")
        body.Append("<tr><td style=\"padding:28px 32px 8px 32px;\">");
        body.Append($"<p style=\"margin:0 0 4px 0; color:{textColor}; font-size:16px; font-weight:600;\">Gentile {WebUtility.HtmlEncode(volunteerSurname)},</p>");
        body.Append($"<p style=\"margin:0; color:{mutedColor}; font-size:14px;\">Ecco i servizi a te assegnati:</p>");
        body.Append("</td></tr>");

        // Assigned rows, each as its own styled card
        body.Append("<tr><td style=\"padding:16px 32px 8px 32px;\">");

        for (int i = 0; i < assignedRows.Count; i++)
        {
            var row = assignedRows[i];

            body.Append($"<table role=\"presentation\" width=\"100%\" cellpadding=\"0\" cellspacing=\"0\" style=\"border:1px solid {cardBorder}; border-left:4px solid {primaryColor}; border-radius:10px; margin-bottom:16px;\">");
            body.Append("<tr><td style=\"padding:16px 18px;\">");
            body.Append($"<div style=\"color:{primaryDark}; font-size:14px; font-weight:700; margin-bottom:10px; text-transform:uppercase; letter-spacing:0.5px;\">Servizio {i + 1}:</div>");

            foreach (var column in row)
            {
                // Skip excluded columns and empty values
                if (!excludedColumns.Contains(column.Key) && !string.IsNullOrWhiteSpace(column.Value))
                {
                    // "Key: Value" stays as one unbroken text run for readability
                    // (and so it can still be matched/searched as plain text).
                    body.Append($"<div style=\"padding:3px 0; color:{textColor}; font-size:13.5px;\">{WebUtility.HtmlEncode(column.Key)}: {WebUtility.HtmlEncode(column.Value)}</div>");
                }
            }

            body.Append("</td></tr>");
            body.Append("</table>");
        }

        body.Append("</td></tr>");

        // Footer
        body.Append("<tr><td style=\"padding:8px 32px 28px 32px;\">");
        body.Append($"<p style=\"margin:0; color:{textColor}; font-size:14px;\">Grazie per la tua disponibilità.</p>");
        body.Append($"<p style=\"margin:16px 0 0 0; color:{mutedColor}; font-size:12px;\">Questo messaggio è stato generato automaticamente da Auser.</p>");
        body.Append("</td></tr>");

        body.Append("</table>");
        body.Append("</td></tr>");
        body.Append("</table>");
        body.Append("</body></html>");

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
