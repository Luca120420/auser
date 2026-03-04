using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Models;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for VolunteerNotificationController email sending functionality using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// </summary>
    [TestFixture]
    public class VolunteerNotificationControllerPropertyTests
    {
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

            // Generate a row as a dictionary
            var rowGen = Gen.Constant(columnNames.Zip(columnValues, (k, v) => (k, v)).ToDictionary(x => x.k, x => x.v));

            return Gen.Choose(1, 10)
                .SelectMany(count => Gen.ListOf(count, rowGen))
                .Select(fsharpList => fsharpList.ToList());
        }

        /// <summary>
        /// Custom generator for volunteer assignments
        /// </summary>
        private static Gen<List<VolunteerAssignment>> VolunteerAssignmentsGen()
        {
            var assignmentGen = from surname in ValidSurnameGen()
                               from email in ValidEmailGen()
                               from assignedRows in AssignedRowsGen()
                               select new VolunteerAssignment
                               {
                                   Surname = surname,
                                   Email = email,
                                   AssignedRows = assignedRows
                               };

            return Gen.Choose(1, 20)
                .SelectMany(count => Gen.ListOf(count, assignmentGen))
                .Select(fsharpList => fsharpList.ToList())
                // Ensure unique surnames to avoid dictionary key conflicts
                .Select(list => list.GroupBy(a => a.Surname)
                                   .Select(g => g.First())
                                   .ToList());
        }

        // Feature: volunteer-email-notifications, Property 10: One Email Per Volunteer With Assignments
        /// <summary>
        /// Property 10: One Email Per Volunteer With Assignments
        /// For any set of volunteer assignments where N volunteers have at least one assigned row,
        /// initiating email sending should result in exactly N email send attempts.
        /// **Validates: Requirements 5.2**
        /// </summary>
        [Test]
        public void Property_OneEmailPerVolunteerWithAssignments()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerAssignmentsGen()),
                (List<VolunteerAssignment> assignments) =>
                {
                    try
                    {
                        // Arrange - Create mock services
                        var mockEmailService = new MockEmailService();
                        var mockVolunteerManager = new MockVolunteerManager();
                        var mockConfigService = new MockConfigurationService();
                        var mockExcelManager = new MockExcelManager(assignments);
                        var mockUI = new MockVolunteerUI();

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Set up prerequisites
                        var volunteers = assignments.ToDictionary(a => a.Surname, a => a.Email);
                        mockVolunteerManager.SetVolunteers(volunteers);
                        mockConfigService.SetGmailCredentials("test@gmail.com", "test-password");
                        
                        controller.OnVolunteerFileSelected("test.json");
                        controller.OnGmailCredentialsUpdated("test@gmail.com", "test-password");
                        controller.OnNotificationExcelFileSelected("test.xlsx");
                        controller.OnSheetSelected("Sheet1");

                        // Act - Send emails
                        var task = controller.OnSendEmailsClickedAsync();
                        task.Wait();

                        // Assert - Count volunteers with at least one assigned row
                        int expectedEmailCount = assignments.Count(a => a.AssignedRows != null && a.AssignedRows.Count > 0);
                        int actualEmailCount = mockEmailService.EmailsSent.Count;

                        if (actualEmailCount != expectedEmailCount)
                        {
                            return false.Label($"Expected {expectedEmailCount} email send attempts, got {actualEmailCount}");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"One email per volunteer test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 13: Email Sending Resilience
        /// <summary>
        /// Property 13: Email Sending Resilience
        /// For any list of volunteers where some email sends fail, the application should
        /// continue attempting to send to all remaining volunteers and not stop at the first failure.
        /// **Validates: Requirements 5.6**
        /// </summary>
        [Test]
        public void Property_EmailSendingResilience()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerAssignmentsGen()),
                (List<VolunteerAssignment> assignments) =>
                {
                    // Skip if we have fewer than 2 assignments (need at least 2 to test resilience)
                    if (assignments.Count < 2)
                    {
                        return true.ToProperty().Label("Skipped: need at least 2 assignments");
                    }

                    try
                    {
                        // Arrange - Create mock services with some failures
                        var mockEmailService = new MockEmailService();
                        
                        // Configure to fail on every other volunteer
                        for (int i = 0; i < assignments.Count; i++)
                        {
                            if (i % 2 == 0)
                            {
                                mockEmailService.SetFailureForEmail(assignments[i].Email);
                            }
                        }

                        var mockVolunteerManager = new MockVolunteerManager();
                        var mockConfigService = new MockConfigurationService();
                        var mockExcelManager = new MockExcelManager(assignments);
                        var mockUI = new MockVolunteerUI();

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Set up prerequisites
                        var volunteers = assignments.ToDictionary(a => a.Surname, a => a.Email);
                        mockVolunteerManager.SetVolunteers(volunteers);
                        mockConfigService.SetGmailCredentials("test@gmail.com", "test-password");
                        
                        controller.OnVolunteerFileSelected("test.json");
                        controller.OnGmailCredentialsUpdated("test@gmail.com", "test-password");
                        controller.OnNotificationExcelFileSelected("test.xlsx");
                        controller.OnSheetSelected("Sheet1");

                        // Act - Send emails
                        var task = controller.OnSendEmailsClickedAsync();
                        task.Wait();

                        // Assert - Verify all volunteers were attempted (not stopped at first failure)
                        int expectedAttempts = assignments.Count(a => a.AssignedRows != null && a.AssignedRows.Count > 0);
                        int actualAttempts = mockEmailService.EmailsSent.Count + mockEmailService.EmailsFailed.Count;

                        if (actualAttempts != expectedAttempts)
                        {
                            return false.Label($"Expected {expectedAttempts} send attempts (including failures), got {actualAttempts}. Application stopped early.");
                        }

                        // Verify that both successes and failures occurred
                        if (mockEmailService.EmailsFailed.Count == 0)
                        {
                            return false.Label("No failures occurred, cannot verify resilience");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Email sending resilience test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 14: Email Summary Accuracy
        /// <summary>
        /// Property 14: Email Summary Accuracy
        /// For any email sending operation with S successful sends and F failed sends,
        /// the displayed summary should report exactly S successes and F failures.
        /// **Validates: Requirements 5.7**
        /// </summary>
        [Test]
        public void Property_EmailSummaryAccuracy()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerAssignmentsGen()),
                (List<VolunteerAssignment> assignments) =>
                {
                    try
                    {
                        // Arrange - Create mock services with random failures
                        var mockEmailService = new MockEmailService();
                        var random = new System.Random();
                        
                        // Randomly fail some emails
                        foreach (var assignment in assignments)
                        {
                            if (random.Next(2) == 0) // 50% chance of failure
                            {
                                mockEmailService.SetFailureForEmail(assignment.Email);
                            }
                        }

                        var mockVolunteerManager = new MockVolunteerManager();
                        var mockConfigService = new MockConfigurationService();
                        var mockExcelManager = new MockExcelManager(assignments);
                        var mockUI = new MockVolunteerUI();

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Set up prerequisites
                        var volunteers = assignments.ToDictionary(a => a.Surname, a => a.Email);
                        mockVolunteerManager.SetVolunteers(volunteers);
                        mockConfigService.SetGmailCredentials("test@gmail.com", "test-password");
                        
                        controller.OnVolunteerFileSelected("test.json");
                        controller.OnGmailCredentialsUpdated("test@gmail.com", "test-password");
                        controller.OnNotificationExcelFileSelected("test.xlsx");
                        controller.OnSheetSelected("Sheet1");

                        // Act - Send emails
                        var task = controller.OnSendEmailsClickedAsync();
                        task.Wait();

                        // Assert - Verify summary accuracy
                        int expectedSuccesses = mockEmailService.EmailsSent.Count;
                        int expectedFailures = mockEmailService.EmailsFailed.Count;

                        int actualSuccesses = mockUI.LastSuccessCount;
                        int actualFailures = mockUI.LastFailureCount;

                        if (actualSuccesses != expectedSuccesses)
                        {
                            return false.Label($"Success count mismatch: expected {expectedSuccesses}, got {actualSuccesses}");
                        }

                        if (actualFailures != expectedFailures)
                        {
                            return false.Label($"Failure count mismatch: expected {expectedFailures}, got {actualFailures}");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Email summary accuracy test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
        }

        #region Mock Implementations

        /// <summary>
        /// Mock implementation of IEmailService for testing
        /// </summary>
        private class MockEmailService : IEmailService
        {
            public List<string> EmailsSent { get; } = new List<string>();
            public List<string> EmailsFailed { get; } = new List<string>();
            private HashSet<string> _failureEmails = new HashSet<string>();

            public void SetFailureForEmail(string email)
            {
                _failureEmails.Add(email);
            }

            public Task<bool> SendVolunteerNotificationAsync(
                string toEmail,
                string volunteerSurname,
                List<Dictionary<string, string>> assignedRows,
                GmailCredentials credentials)
            {
                if (_failureEmails.Contains(toEmail))
                {
                    EmailsFailed.Add(toEmail);
                    return Task.FromResult(false);
                }

                EmailsSent.Add(toEmail);
                return Task.FromResult(true);
            }

            public string FormatEmailBody(string volunteerSurname, List<Dictionary<string, string>> assignedRows)
            {
                return $"Email body for {volunteerSurname}";
            }

            public Task<bool> TestConnectionAsync(GmailCredentials credentials)
            {
                return Task.FromResult(true);
            }
        }

        /// <summary>
        /// Mock implementation of IVolunteerManager for testing
        /// </summary>
        private class MockVolunteerManager : IVolunteerManager
        {
            private Dictionary<string, string> _volunteers = new Dictionary<string, string>();

            public void SetVolunteers(Dictionary<string, string> volunteers)
            {
                _volunteers = new Dictionary<string, string>(volunteers);
            }

            public Dictionary<string, string> LoadVolunteers(string filePath)
            {
                return new Dictionary<string, string>(_volunteers);
            }

            public void SaveVolunteers(string filePath, Dictionary<string, string> volunteers)
            {
                // No-op for testing
            }

            public void AddVolunteer(string surname, string email, Dictionary<string, string> volunteers)
            {
                volunteers[surname] = email;
            }

            public void RemoveVolunteer(string surname, Dictionary<string, string> volunteers)
            {
                volunteers.Remove(surname);
            }

            public bool IsValidEmail(string email)
            {
                return !string.IsNullOrWhiteSpace(email) && email.Contains("@");
            }
        }

        /// <summary>
        /// Mock implementation of IConfigurationService for testing
        /// </summary>
        private class MockConfigurationService : IConfigurationService
        {
            private AppConfiguration _config = new AppConfiguration();

            public void SetGmailCredentials(string email, string password)
            {
                _config.GmailCredentials.Email = email;
                _config.GmailCredentials.AppPassword = password;
            }

            public AppConfiguration LoadConfiguration()
            {
                return _config;
            }

            public void SaveConfiguration(AppConfiguration config)
            {
                _config = config;
            }

            public string GetConfigFilePath()
            {
                return "test-config.json";
            }
        }

        /// <summary>
        /// Mock implementation of IExcelManager for testing
        /// </summary>
        private class MockExcelManager : IExcelManager
        {
            private List<VolunteerAssignment> _assignments;

            public MockExcelManager(List<VolunteerAssignment> assignments)
            {
                _assignments = assignments;
            }

            public ExcelWorkbook OpenWorkbook(string filePath)
            {
                // Create a minimal EPPlus package for testing
                var package = new OfficeOpenXml.ExcelPackage();
                return new ExcelWorkbook(package);
            }

            public List<string> GetSheetNames(ExcelWorkbook workbook)
            {
                return new List<string> { "Sheet1" };
            }

            public Sheet GetSheetByName(ExcelWorkbook workbook, string sheetName)
            {
                // Create a minimal worksheet for testing
                var package = new OfficeOpenXml.ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                return new Sheet(worksheet);
            }

            public List<VolunteerAssignment> IdentifyVolunteerAssignments(Sheet sheet, Dictionary<string, string> volunteers)
            {
                return _assignments;
            }

            // Other methods not used in these tests
            public int GetNextSheetNumber(List<string> sheetNames) => 1;
            public Sheet GetFissiSheet(ExcelWorkbook workbook) 
            {
                var package = new OfficeOpenXml.ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("fissi");
                return new Sheet(worksheet);
            }
            public Sheet CreateNewSheet(ExcelWorkbook workbook, int sheetNumber) 
            {
                var package = new OfficeOpenXml.ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add($"Sheet{sheetNumber}");
                return new Sheet(worksheet);
            }
            public void WriteHeaderRow(Sheet sheet, DateTime mondayDate) { }
            public void WriteColumnHeaders(Sheet sheet) { }
            public void WriteDataRows(Sheet sheet, List<TransformedRow> rows, int startRow) { }
            public void AppendFissiData(Sheet targetSheet, Sheet fissiSheet, int startRow) { }
            public void ApplyYellowHighlight(Sheet sheet, List<int> rowNumbers) { }
            public void EnableAutoFilter(Sheet sheet) { }
            public void SaveWorkbook(ExcelWorkbook workbook, string filePath) { }
            public string ReadHeader(Sheet sheet) => "";
            public void WriteColumnHeadersEnhanced(Sheet sheet) { }
            public void WriteDataRowsEnhanced(Sheet sheet, List<EnhancedTransformedRow> rows, int startRow) { }
            public void SortDataRows(Sheet sheet, int startRow, int endRow) { }
            public void ApplyBoldToHeaders(Sheet sheet, int headerRow) { }
            public void ApplyThickBordersToDateGroups(Sheet sheet, int startRow, int endRow) { }
        }

        /// <summary>
        /// Mock implementation of IVolunteerUI for testing
        /// </summary>
        private class MockVolunteerUI : IVolunteerUI
        {
            public int LastSuccessCount { get; private set; }
            public int LastFailureCount { get; private set; }

            public void DisplayVolunteerList(Dictionary<string, string> volunteers) { }
            public void DisplayGmailCredentials(string email, string password) { }
            public void DisplaySheetNames(List<string> sheetNames) { }
            public void EnableSendEmailsButton(bool enabled) { }
            public void ShowEmailProgress(string message) { }
            
            public void ShowEmailSummary(int successCount, int failureCount)
            {
                LastSuccessCount = successCount;
                LastFailureCount = failureCount;
            }

            public bool ConfirmAction(string message) => true;
            public void ShowErrorMessage(string message) { }
        }

        #endregion
    }
}
