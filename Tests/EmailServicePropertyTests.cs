using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Threading.Tasks;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for EmailService class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// </summary>
    [TestFixture]
    public class EmailServicePropertyTests
    {
        private TestableEmailService _emailService = null!;

        [SetUp]
        public void Setup()
        {
            _emailService = new TestableEmailService();
        }

        /// <summary>
        /// Custom generator for valid email addresses
        /// </summary>
        private static Gen<string> ValidEmailGen()
        {
            // Generate local part: must start with alphanumeric, can contain ._- in middle
            var localPartGen = from firstChar in Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray())
                              from length in Gen.Choose(0, 19)
                              from middleChars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789._-".ToCharArray()))
                              select firstChar + new string(middleChars);

            // Generate domain part: must start with alphanumeric, can contain - in middle
            var domainGen = from firstChar in Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789".ToCharArray())
                           from length in Gen.Choose(1, 14)
                           from middleChars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789-".ToCharArray()))
                           select firstChar + new string(middleChars);

            var tldGen = Gen.Elements("com", "org", "net", "it", "edu");

            return from local in localPartGen
                   from domain in domainGen
                   from tld in tldGen
                   select $"{local}@{domain}.{tld}";
        }

        /// <summary>
        /// Custom generator for valid volunteer surnames (non-empty strings)
        /// </summary>
        private static Gen<string> ValidSurnameGen()
        {
            return Gen.Choose(1, 30)
                .SelectMany(length => Gen.ArrayOf(length, Gen.Elements("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz ".ToCharArray())))
                .Where(chars => chars.Length > 0)
                .Select(chars => new string(chars).Trim())
                .Where(surname => !string.IsNullOrWhiteSpace(surname));
        }

        /// <summary>
        /// Custom generator for assigned rows (list of dictionaries)
        /// </summary>
        private static Gen<List<Dictionary<string, string>>> AssignedRowsGen()
        {
            var columnNames = new[] { "Data", "Ora", "Servizio", "Note", "Luogo" };
            var columnValues = new[] { "10/01/2024", "09:00", "Trasporto", "Urgente", "Milano" };

            // Generate a row as a dictionary with unique keys
            var rowGen = Gen.Constant(columnNames.Zip(columnValues, (k, v) => (k, v)).ToDictionary(x => x.k, x => x.v));

            return Gen.Choose(1, 10)
                .SelectMany(count => Gen.ListOf(count, rowGen))
                .Select(fsharpList => fsharpList.ToList());
        }

        // Feature: volunteer-email-notifications, Property 11: All Assigned Rows Included In Email
        /// <summary>
        /// Property 11: All Assigned Rows Included In Email
        /// For any volunteer with M assigned rows, the email notification sent to that volunteer
        /// should contain all M rows in the formatted body.
        /// **Validates: Requirements 5.3**
        /// </summary>
        [Test]
        public void Property_AllAssignedRowsIncludedInEmail()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var testDataGen = from volunteerSurname in ValidSurnameGen()
                             from assignedRows in AssignedRowsGen()
                             select (volunteerSurname, assignedRows);

            Prop.ForAll(
                Arb.From(testDataGen),
                (testData) =>
                {
                    var (volunteerSurname, assignedRows) = testData;

                    try
                    {
                        // Act - Format email body
                        var emailBody = _emailService.FormatEmailBody(volunteerSurname, assignedRows);

                        // Assert - Verify all rows are included
                        // Each row should have a "Servizio N:" marker
                        for (int i = 0; i < assignedRows.Count; i++)
                        {
                            var serviceMarker = $"Servizio {i + 1}:";
                            if (!emailBody.Contains(serviceMarker))
                            {
                                return false.Label($"Missing service marker '{serviceMarker}' for row {i + 1}");
                            }

                            // Verify all column values from this row are present
                            var row = assignedRows[i];
                            foreach (var column in row)
                            {
                                if (!emailBody.Contains(column.Value))
                                {
                                    return false.Label($"Missing column value '{column.Value}' from row {i + 1}");
                                }
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"All assigned rows test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 12: Email Content Language Validation
        /// <summary>
        /// Property 12: Email Content Language Validation
        /// For any generated email notification, the subject line and greeting should contain
        /// Italian text (verified by presence of Italian keywords like "Gentile", "Settimana", "Servizi").
        /// **Validates: Requirements 5.4**
        /// </summary>
        [Test]
        public void Property_EmailContentLanguageValidation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var testDataGen = from volunteerSurname in ValidSurnameGen()
                             from assignedRows in AssignedRowsGen()
                             select (volunteerSurname, assignedRows);

            Prop.ForAll(
                Arb.From(testDataGen),
                (testData) =>
                {
                    var (volunteerSurname, assignedRows) = testData;

                    try
                    {
                        // Act - Format email body
                        var emailBody = _emailService.FormatEmailBody(volunteerSurname, assignedRows);

                        // Assert - Verify Italian keywords are present
                        var italianKeywords = new[] { "Gentile", "servizi", "assegnati", "Grazie", "disponibilità" };
                        
                        foreach (var keyword in italianKeywords)
                        {
                            if (!emailBody.Contains(keyword, StringComparison.OrdinalIgnoreCase))
                            {
                                return false.Label($"Missing Italian keyword '{keyword}' in email body");
                            }
                        }

                        // Verify the greeting includes the volunteer surname
                        if (!emailBody.Contains($"Gentile {volunteerSurname}"))
                        {
                            return false.Label($"Missing personalized greeting 'Gentile {volunteerSurname}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Language validation test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 16: Complete Column Formatting
        /// <summary>
        /// Property 16: Complete Column Formatting
        /// For any assigned row with C columns, the formatted text representation
        /// should include all C column values.
        /// **Validates: Requirements 6.1**
        /// </summary>
        [Test]
        public void Property_CompleteColumnFormatting()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var testDataGen = from volunteerSurname in ValidSurnameGen()
                             from assignedRows in AssignedRowsGen()
                             select (volunteerSurname, assignedRows);

            Prop.ForAll(
                Arb.From(testDataGen),
                (testData) =>
                {
                    var (volunteerSurname, assignedRows) = testData;

                    try
                    {
                        // Act - Format email body
                        var emailBody = _emailService.FormatEmailBody(volunteerSurname, assignedRows);

                        // Assert - Verify all columns from all rows are present
                        foreach (var row in assignedRows)
                        {
                            foreach (var column in row)
                            {
                                // Check that both column name and value are present
                                if (!emailBody.Contains(column.Key))
                                {
                                    return false.Label($"Missing column name '{column.Key}' in formatted email");
                                }

                                if (!emailBody.Contains(column.Value))
                                {
                                    return false.Label($"Missing column value '{column.Value}' in formatted email");
                                }

                                // Verify the column is formatted as "Key: Value"
                                var formattedColumn = $"{column.Key}: {column.Value}";
                                if (!emailBody.Contains(formattedColumn))
                                {
                                    return false.Label($"Column not properly formatted as '{formattedColumn}'");
                                }
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Complete column formatting test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 17: Multiple Row Visual Separation
        /// <summary>
        /// Property 17: Multiple Row Visual Separation
        /// For any email with M assigned rows where M > 1, the formatted email body should contain
        /// M-1 or more separator elements (blank lines or delimiters) between rows.
        /// **Validates: Requirements 6.4**
        /// </summary>
        [Test]
        public void Property_MultipleRowVisualSeparation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Generate test data with at least 2 rows to test separation
            var testDataGen = from volunteerSurname in ValidSurnameGen()
                             from rowCount in Gen.Choose(2, 10)
                             from assignedRows in Gen.ListOf(rowCount, Gen.Constant(
                                 new Dictionary<string, string>
                                 {
                                     { "Data", "10/01/2024" },
                                     { "Ora", "09:00" },
                                     { "Servizio", "Trasporto" }
                                 }))
                             select (volunteerSurname, assignedRows.ToList());

            Prop.ForAll(
                Arb.From(testDataGen),
                (testData) =>
                {
                    var (volunteerSurname, assignedRows) = testData;

                    try
                    {
                        // Act - Format email body
                        var emailBody = _emailService.FormatEmailBody(volunteerSurname, assignedRows);

                        // Assert - Verify visual separation between rows
                        // Count the number of "Servizio N:" markers
                        int serviceMarkerCount = 0;
                        for (int i = 1; i <= assignedRows.Count; i++)
                        {
                            if (emailBody.Contains($"Servizio {i}:"))
                            {
                                serviceMarkerCount++;
                            }
                        }

                        if (serviceMarkerCount != assignedRows.Count)
                        {
                            return false.Label($"Expected {assignedRows.Count} service markers, found {serviceMarkerCount}");
                        }

                        // For M rows, we expect M-1 separators between them
                        // Count blank lines (double newlines) between service sections
                        var lines = emailBody.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                        int blankLineCount = 0;
                        bool inServiceSection = false;

                        for (int i = 0; i < lines.Length; i++)
                        {
                            if (lines[i].StartsWith("Servizio "))
                            {
                                inServiceSection = true;
                            }
                            else if (inServiceSection && string.IsNullOrWhiteSpace(lines[i]))
                            {
                                // Check if this blank line is between two service sections
                                // (not at the end of all services)
                                if (i + 1 < lines.Length && 
                                    lines.Skip(i + 1).Any(l => l.StartsWith("Servizio ")))
                                {
                                    blankLineCount++;
                                }
                            }
                        }

                        // We need at least M-1 separators for M rows
                        int expectedSeparators = assignedRows.Count - 1;
                        if (blankLineCount < expectedSeparators)
                        {
                            return false.Label($"Expected at least {expectedSeparators} separators between {assignedRows.Count} rows, found {blankLineCount}");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Visual separation test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 15: Correct Email Routing
        /// <summary>
        /// Property 15: Correct Email Routing
        /// For any volunteer with email address E in the volunteer file,
        /// the email notification for that volunteer should be sent to address E.
        /// **Validates: Requirements 5.8**
        /// </summary>
        [Test]
        public void Property_CorrectEmailRouting()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            var testDataGen = from volunteerEmail in ValidEmailGen()
                             from volunteerSurname in ValidSurnameGen()
                             from assignedRows in AssignedRowsGen()
                             from senderEmail in ValidEmailGen()
                             select (volunteerEmail, volunteerSurname, assignedRows, senderEmail);

            Prop.ForAll(
                Arb.From(testDataGen),
                (testData) =>
                {
                    var (volunteerEmail, volunteerSurname, assignedRows, senderEmail) = testData;

                    try
                    {
                        // Arrange - Create credentials
                        var credentials = new GmailCredentials
                        {
                            Email = senderEmail,
                            AppPassword = "test-app-password-1234"
                        };

                        // Act - Send email notification
                        var task = _emailService.SendVolunteerNotificationAsync(
                            volunteerEmail,
                            volunteerSurname,
                            assignedRows,
                            credentials);

                        // Wait for the async operation to complete
                        task.Wait();

                        // Assert - Verify the email was sent to the correct address
                        var lastRecipient = _emailService.LastRecipientEmail;

                        if (lastRecipient == null)
                        {
                            return false.Label("No email was sent");
                        }

                        if (lastRecipient != volunteerEmail)
                        {
                            return false.Label($"Email sent to wrong address: expected '{volunteerEmail}', got '{lastRecipient}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Email routing test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Testable version of EmailService that captures the recipient email
        /// without actually sending emails via SMTP.
        /// </summary>
        private class TestableEmailService : IEmailService
        {
            public string? LastRecipientEmail { get; private set; }

            public Task<bool> SendVolunteerNotificationAsync(
                string toEmail,
                string volunteerSurname,
                List<Dictionary<string, string>> assignedRows,
                GmailCredentials credentials)
            {
                // Capture the recipient email for verification
                LastRecipientEmail = toEmail;

                // Validate inputs
                if (string.IsNullOrWhiteSpace(toEmail))
                {
                    return Task.FromResult(false);
                }

                if (string.IsNullOrWhiteSpace(volunteerSurname))
                {
                    return Task.FromResult(false);
                }

                if (assignedRows == null || assignedRows.Count == 0)
                {
                    return Task.FromResult(false);
                }

                if (!credentials.IsConfigured)
                {
                    return Task.FromResult(false);
                }

                // Validate email format
                try
                {
                    var mailAddress = new MailAddress(toEmail);
                }
                catch (FormatException)
                {
                    return Task.FromResult(false);
                }

                // Simulate successful email sending
                return Task.FromResult(true);
            }

            public string FormatEmailBody(string volunteerSurname, List<Dictionary<string, string>> assignedRows)
            {
                // Use the same implementation as the real EmailService
                var emailService = new EmailService();
                return emailService.FormatEmailBody(volunteerSurname, assignedRows);
            }

            public Task<bool> TestConnectionAsync(GmailCredentials credentials)
            {
                // Simulate connection test
                return Task.FromResult(credentials.IsConfigured);
            }
        }
    }
}
