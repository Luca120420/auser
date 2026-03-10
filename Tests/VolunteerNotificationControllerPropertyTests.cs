using System;
using System.Collections.Generic;
using System.IO;
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

        /// <summary>
        /// Custom generator for volunteer dictionaries (surname -> email mappings)
        /// </summary>
        private static Gen<Dictionary<string, string>> VolunteerDictionaryGen()
        {
            var volunteerPairGen = from surname in ValidSurnameGen()
                                  from email in ValidEmailGen()
                                  select (surname, email);

            return Gen.Choose(0, 15)
                .SelectMany(count => Gen.ListOf(count, volunteerPairGen))
                .Select(fsharpList => fsharpList.ToList())
                // Ensure unique surnames
                .Select(list => list.GroupBy(v => v.surname)
                                   .Select(g => g.First())
                                   .ToDictionary(v => v.surname, v => v.email));
        }

        // Feature: portable-data-storage, Property 4: Volunteer Import Merges Data
        /// <summary>
        /// Property 4: Volunteer Import Merges Data
        /// For any existing volunteer dictionary and any imported volunteer dictionary,
        /// importing should result in a merged dictionary containing all surnames from both sources,
        /// with imported email addresses overwriting existing ones for duplicate surnames.
        /// **Validates: Requirements 3.2, 3.3**
        /// </summary>
        [Test]
        public void Property_VolunteerImportMergesData()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> existingVolunteers, Dictionary<string, string> importedVolunteers) =>
                {
                    try
                    {
                        // Arrange - Create mock services
                        var mockVolunteerManager = new MockVolunteerManager();
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationService();
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Set up existing volunteers
                        mockVolunteerManager.SetVolunteers(existingVolunteers);

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Load existing volunteers
                        if (existingVolunteers.Count > 0)
                        {
                            controller.OnVolunteerFileSelected("existing.json");
                        }

                        // Calculate expected merged result
                        var expectedMerged = new Dictionary<string, string>(existingVolunteers);
                        foreach (var imported in importedVolunteers)
                        {
                            expectedMerged[imported.Key] = imported.Value; // Add new or overwrite existing
                        }

                        // Act - Import volunteers (this should merge with existing)
                        mockVolunteerManager.SetVolunteers(importedVolunteers);
                        controller.OnVolunteerFileSelected("imported.json");

                        // Assert - Get actual merged volunteers
                        var actualMerged = controller.GetVolunteers();

                        // Verify all expected surnames are present
                        if (actualMerged.Count != expectedMerged.Count)
                        {
                            return false.Label($"Expected {expectedMerged.Count} volunteers after merge, got {actualMerged.Count}");
                        }

                        // Verify each surname has the correct email
                        foreach (var expected in expectedMerged)
                        {
                            if (!actualMerged.ContainsKey(expected.Key))
                            {
                                return false.Label($"Missing surname '{expected.Key}' in merged volunteers");
                            }

                            if (actualMerged[expected.Key] != expected.Value)
                            {
                                return false.Label($"Email mismatch for surname '{expected.Key}': expected '{expected.Value}', got '{actualMerged[expected.Key]}'");
                            }
                        }

                        // Verify no extra surnames
                        foreach (var actual in actualMerged)
                        {
                            if (!expectedMerged.ContainsKey(actual.Key))
                            {
                                return false.Label($"Unexpected surname '{actual.Key}' in merged volunteers");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Volunteer import merge test failed with exception: {ex.Message}");
                    }
                }
            ).Check(config);
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

        // Feature: portable-data-storage, Property 5: Import Persists to Internal Storage
        /// <summary>
        /// Property 5: Import Persists to Internal Storage
        /// For any external volunteer file, importing it should result in the volunteer data being
        /// saved to data/volunteers.json, and the data should be loadable from that location.
        /// **Validates: Requirements 3.1, 3.4**
        /// </summary>
        [Test]
        public void Property_ImportPersistsToInternalStorage()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> importedVolunteers) =>
                {
                    // Skip empty dictionaries
                    if (importedVolunteers.Count == 0)
                    {
                        return true.ToProperty().Label("Skipped: empty volunteer dictionary");
                    }

                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");
                    var volunteersPath = Path.Combine(dataFolder, "volunteers.json");
                    var externalFilePath = Path.Combine(Path.GetTempPath(), $"external_{Guid.NewGuid()}.json");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder); // Ensure data folder exists

                        // Create external volunteer file
                        var externalData = new { associates = importedVolunteers };
                        var externalJson = System.Text.Json.JsonSerializer.Serialize(externalData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(externalFilePath, externalJson);

                        // Create mock services with custom volunteer manager that uses temp directory
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Ensure data folder exists via config service
                        mockConfigService.EnsureDataFolderExists();

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Act - Import volunteers from external file
                        controller.OnVolunteerFileSelected(externalFilePath);

                        // Assert - Verify data was saved to internal storage
                        if (!File.Exists(volunteersPath))
                        {
                            return false.Label($"Volunteer data was not saved to internal storage at: {volunteersPath}");
                        }

                        // Assert - Verify data can be loaded from internal storage
                        var loadedVolunteers = mockVolunteerManager.LoadVolunteers(volunteersPath);

                        if (loadedVolunteers.Count != importedVolunteers.Count)
                        {
                            return false.Label($"Loaded volunteer count mismatch: expected {importedVolunteers.Count}, got {loadedVolunteers.Count}");
                        }

                        // Verify each volunteer was persisted correctly
                        foreach (var imported in importedVolunteers)
                        {
                            if (!loadedVolunteers.ContainsKey(imported.Key))
                            {
                                return false.Label($"Missing surname '{imported.Key}' in persisted data");
                            }

                            if (loadedVolunteers[imported.Key] != imported.Value)
                            {
                                return false.Label($"Email mismatch for surname '{imported.Key}': expected '{imported.Value}', got '{loadedVolunteers[imported.Key]}'");
                            }
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Import persistence test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempAppDir))
                        {
                            try { Directory.Delete(tempAppDir, true); } catch { }
                        }
                        if (File.Exists(externalFilePath))
                        {
                            try { File.Delete(externalFilePath); } catch { }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 6: No External Volunteer Paths Stored
        /// <summary>
        /// Property 6: No External Volunteer Paths Stored
        /// For any volunteer import operation, after completion, the configuration file should not
        /// contain any reference to the external volunteer file path.
        /// **Validates: Requirements 2.4, 3.5, 4.3**
        /// </summary>
        [Test]
        public void Property_NoExternalVolunteerPathsStored()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> importedVolunteers) =>
                {
                    // Skip empty dictionaries
                    if (importedVolunteers.Count == 0)
                    {
                        return true.ToProperty().Label("Skipped: empty volunteer dictionary");
                    }

                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");
                    var configPath = Path.Combine(dataFolder, "config.json");
                    var externalFilePath = Path.Combine(Path.GetTempPath(), $"external_{Guid.NewGuid()}.json");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create external volunteer file
                        var externalData = new { associates = importedVolunteers };
                        var externalJson = System.Text.Json.JsonSerializer.Serialize(externalData, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                        File.WriteAllText(externalFilePath, externalJson);

                        // Create mock services with real file I/O
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Create controller
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI);

                        // Act - Import volunteers from external file
                        controller.OnVolunteerFileSelected(externalFilePath);

                        // Assert - Verify configuration file exists
                        if (!File.Exists(configPath))
                        {
                            // If config doesn't exist, that's fine - no external path is stored
                            return true.ToProperty();
                        }

                        // Assert - Read configuration file and verify no external path is stored
                        var configJson = File.ReadAllText(configPath);
                        
                        // Check that the external file path is NOT in the configuration
                        if (configJson.Contains(externalFilePath))
                        {
                            return false.Label($"Configuration contains external file path: {externalFilePath}");
                        }

                        // Check that no "VolunteerFilePath" property exists in the configuration
                        if (configJson.Contains("VolunteerFilePath"))
                        {
                            return false.Label("Configuration contains 'VolunteerFilePath' property, which should not exist");
                        }

                        // Parse the configuration to verify structure
                        var configDoc = System.Text.Json.JsonDocument.Parse(configJson);
                        var root = configDoc.RootElement;

                        // Verify VolunteerFilePath property does not exist
                        if (root.TryGetProperty("VolunteerFilePath", out _))
                        {
                            return false.Label("Configuration JSON contains 'VolunteerFilePath' property");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"No external paths stored test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempAppDir))
                        {
                            try { Directory.Delete(tempAppDir, true); } catch { }
                        }
                        if (File.Exists(externalFilePath))
                        {
                            try { File.Delete(externalFilePath); } catch { }
                        }
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

            public void EnsureDataFolderExists()
            {
                // Mock implementation - do nothing for tests
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
            public void AppendFissiData(Sheet targetSheet, Sheet fissiSheet, int startRow, DateTime targetWeekMonday) { }
            public void AppendLaboratoriData(Sheet targetSheet, Sheet laboratoriSheet, int startRow, DateTime targetWeekMonday) { }
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

        /// <summary>
        /// Mock implementation of IVolunteerManager with real file I/O for testing persistence
        /// </summary>
        private class MockVolunteerManagerWithRealIO : IVolunteerManager
        {
            private readonly string _dataFolder;

            public MockVolunteerManagerWithRealIO(string dataFolder)
            {
                _dataFolder = dataFolder;
            }

            public Dictionary<string, string> LoadVolunteers(string filePath)
            {
                if (!File.Exists(filePath))
                {
                    return new Dictionary<string, string>();
                }

                var json = File.ReadAllText(filePath);
                var data = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, string>>>(json);
                return data?["associates"] ?? new Dictionary<string, string>();
            }

            public void SaveVolunteers(string filePath, Dictionary<string, string> volunteers)
            {
                // Ensure directory exists before saving
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var data = new { associates = volunteers };
                var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
                var json = System.Text.Json.JsonSerializer.Serialize(data, options);
                File.WriteAllText(filePath, json);
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
        /// Mock implementation of IConfigurationService with custom app directory for testing
        /// </summary>
        private class MockConfigurationServiceWithAppDir : IConfigurationService
        {
            private readonly string _appDir;
            private AppConfiguration _config = new AppConfiguration();

            public MockConfigurationServiceWithAppDir(string appDir)
            {
                _appDir = appDir;
            }

            public AppConfiguration LoadConfiguration()
            {
                var configPath = GetConfigFilePath();
                if (File.Exists(configPath))
                {
                    var json = File.ReadAllText(configPath);
                    _config = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json) ?? new AppConfiguration();
                }
                return _config;
            }

            public void SaveConfiguration(AppConfiguration config)
            {
                _config = config;
                EnsureDataFolderExists();
                var configPath = GetConfigFilePath();
                var options = new System.Text.Json.JsonSerializerOptions { WriteIndented = true };
                var json = System.Text.Json.JsonSerializer.Serialize(config, options);
                File.WriteAllText(configPath, json);
            }

            public string GetConfigFilePath()
            {
                return Path.Combine(_appDir, "data", "config.json");
            }

            public void EnsureDataFolderExists()
            {
                var dataFolder = Path.Combine(_appDir, "data");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }
            }
        }

        #endregion
    }
}
