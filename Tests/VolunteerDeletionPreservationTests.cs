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
    /// Preservation property tests for volunteer deletion persistence fix.
    /// These tests verify that non-deletion operations remain unchanged by the fix.
    /// Tests are run on UNFIXED code to establish baseline behavior.
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// </summary>
    [TestFixture]
    public class VolunteerDeletionPreservationTests
    {
        /// <summary>
        /// Custom generator for volunteer dictionaries (surname -> email mappings)
        /// </summary>
        private static Gen<Dictionary<string, string>> VolunteerDictionaryGen()
        {
            var validEmailGen = from local in Gen.Elements("john", "jane", "bob", "alice", "charlie")
                               from domain in Gen.Elements("example", "test", "demo")
                               from tld in Gen.Elements("com", "org", "net", "it")
                               select $"{local}@{domain}.{tld}";

            var validSurnameGen = Gen.Elements("Rossi", "Bianchi", "Verdi", "Neri", "Gialli", "Blu");

            var volunteerPairGen = from surname in validSurnameGen
                                  from email in validEmailGen
                                  select (surname, email);

            return Gen.Choose(2, 5)
                .SelectMany(count => Gen.ListOf(count, volunteerPairGen))
                .Select(fsharpList => fsharpList.ToList())
                .Select(list => list.GroupBy(v => v.surname)
                                   .Select(g => g.First())
                                   .ToDictionary(v => v.surname, v => v.email));
        }

        /// <summary>
        /// Custom generator for valid email addresses
        /// </summary>
        private static Gen<string> ValidEmailGen()
        {
            return from local in Gen.Elements("john", "jane", "bob", "alice", "charlie", "david", "emma")
                   from domain in Gen.Elements("example", "test", "demo", "mail")
                   from tld in Gen.Elements("com", "org", "net", "it", "edu")
                   select $"{local}@{domain}.{tld}";
        }

        /// <summary>
        /// Custom generator for valid surnames
        /// </summary>
        private static Gen<string> ValidSurnameGen()
        {
            return Gen.Elements("Rossi", "Bianchi", "Verdi", "Neri", "Gialli", "Blu", "Viola", "Arancione");
        }

        /// <summary>
        /// Property 2: Preservation - OnAddVolunteer Continues to Work
        /// 
        /// This test captures the ACTUAL behavior of OnAddVolunteer on unfixed code.
        /// The fix should not change this behavior.
        /// 
        /// NOTE: OnAddVolunteer has the same bug (doesn't save to internal storage),
        /// but that's out of scope for this bugfix. We're preserving the current behavior.
        /// 
        /// For any valid surname and email, OnAddVolunteer SHALL:
        /// - Add the volunteer to the in-memory dictionary
        /// - Update the UI display
        /// - Update the send emails button state
        /// 
        /// **Validates: Requirements 3.1, 3.5**
        /// </summary>
        [Test]
        public void Property_OnAddVolunteer_ContinuesToWork()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                Arb.From(ValidSurnameGen()),
                Arb.From(ValidEmailGen()),
                (string newSurname, string newEmail) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create mock services
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Create controller (starts with empty volunteers)
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        var initialCount = controller.GetVolunteers().Count;

                        // Act - Add a new volunteer
                        controller.OnAddVolunteer(newSurname, newEmail);

                        // Assert - Verify volunteer is added to in-memory dictionary
                        var volunteers = controller.GetVolunteers();
                        if (!volunteers.ContainsKey(newSurname))
                        {
                            return false.Label($"OnAddVolunteer failed: volunteer '{newSurname}' not added to in-memory dictionary");
                        }

                        if (volunteers[newSurname] != newEmail)
                        {
                            return false.Label($"OnAddVolunteer failed: email mismatch. Expected '{newEmail}', got '{volunteers[newSurname]}'");
                        }

                        if (volunteers.Count != initialCount + 1)
                        {
                            return false.Label($"OnAddVolunteer failed: incorrect count. Expected {initialCount + 1}, got {volunteers.Count}");
                        }

                        // Assert - Verify UI was updated (DisplayVolunteerList was called)
                        if (!mockUI.DisplayVolunteerListCalled)
                        {
                            return false.Label("OnAddVolunteer failed: UI was not updated (DisplayVolunteerList not called)");
                        }

                        // Assert - Verify send emails button state was updated
                        if (!mockUI.EnableSendEmailsButtonCalled)
                        {
                            return false.Label("OnAddVolunteer failed: send emails button state was not updated");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempAppDir))
                        {
                            try { Directory.Delete(tempAppDir, true); } catch { }
                        }
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 2: Preservation - OnVolunteerFileSelected Continues to Save to Internal Storage
        /// 
        /// This test captures the ACTUAL behavior of OnVolunteerFileSelected on unfixed code.
        /// The fix should not change this behavior.
        /// 
        /// For any imported volunteer file, OnVolunteerFileSelected SHALL:
        /// - Merge imported volunteers with existing ones (imported override existing)
        /// - Save to internal storage (volunteers.json)
        /// - Update the UI display
        /// - Update the send emails button state
        /// 
        /// **Validates: Requirements 3.2, 3.4, 3.5**
        /// </summary>
        [Test]
        public void Property_OnVolunteerFileSelected_ContinuesToSaveToInternalStorage()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> importedVolunteers) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");
                    var volunteersPath = Path.Combine(dataFolder, "volunteers.json");
                    var externalFilePath = Path.Combine(Path.GetTempPath(), $"external_{Guid.NewGuid()}.json");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create mock services
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Save imported volunteers to external file
                        mockVolunteerManager.SaveVolunteers(externalFilePath, importedVolunteers);

                        // Create controller (starts with empty volunteers)
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        // Act - Import volunteers from external file
                        controller.OnVolunteerFileSelected(externalFilePath);

                        // Assert - Verify internal storage is updated with imported volunteers
                        var internalVolunteers = mockVolunteerManager.LoadVolunteers(volunteersPath);
                        
                        if (internalVolunteers.Count != importedVolunteers.Count)
                        {
                            return false.Label($"OnVolunteerFileSelected failed: internal storage count mismatch. Expected {importedVolunteers.Count}, got {internalVolunteers.Count}");
                        }

                        foreach (var kvp in importedVolunteers)
                        {
                            if (!internalVolunteers.ContainsKey(kvp.Key))
                            {
                                return false.Label($"OnVolunteerFileSelected failed: volunteer '{kvp.Key}' missing from internal storage");
                            }
                            if (internalVolunteers[kvp.Key] != kvp.Value)
                            {
                                return false.Label($"OnVolunteerFileSelected failed: email mismatch for '{kvp.Key}'. Expected '{kvp.Value}', got '{internalVolunteers[kvp.Key]}'");
                            }
                        }

                        // Assert - Verify UI was updated
                        if (!mockUI.DisplayVolunteerListCalled)
                        {
                            return false.Label("OnVolunteerFileSelected failed: UI was not updated");
                        }

                        // Assert - Verify send emails button state was updated
                        if (!mockUI.EnableSendEmailsButtonCalled)
                        {
                            return false.Label("OnVolunteerFileSelected failed: send emails button state was not updated");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Test failed with exception: {ex.Message}");
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

        /// <summary>
        /// Property 2: Preservation - Error Handling for File I/O Operations Remains Unchanged
        /// 
        /// This test captures the ACTUAL error handling behavior on unfixed code.
        /// The fix should not change this behavior.
        /// 
        /// For any file I/O error during volunteer operations, the system SHALL:
        /// - Display appropriate error messages to the user
        /// - Not crash or leave the system in an inconsistent state
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Test]
        public void Property_ErrorHandling_RemainsUnchanged()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 30;

            Prop.ForAll(
                Arb.From(Gen.Constant(true)),
                (_) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create mock services
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
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        // Test Case 1: File not found error
                        var nonExistentFile = Path.Combine(Path.GetTempPath(), $"nonexistent_{Guid.NewGuid()}.json");
                        controller.OnVolunteerFileSelected(nonExistentFile);

                        // Assert - Verify error message was shown
                        if (!mockUI.ShowErrorMessageCalled)
                        {
                            return false.Label("Error handling failed: ShowErrorMessage not called for file not found");
                        }

                        // Reset mock UI
                        mockUI.ShowErrorMessageCalled = false;

                        // Test Case 2: Invalid file format error
                        var invalidFile = Path.Combine(Path.GetTempPath(), $"invalid_{Guid.NewGuid()}.json");
                        File.WriteAllText(invalidFile, "{ invalid json content }");
                        controller.OnVolunteerFileSelected(invalidFile);

                        // Assert - Verify error message was shown
                        if (!mockUI.ShowErrorMessageCalled)
                        {
                            File.Delete(invalidFile);
                            return false.Label("Error handling failed: ShowErrorMessage not called for invalid file format");
                        }

                        File.Delete(invalidFile);

                        // Test Case 3: Invalid email validation
                        mockUI.ShowErrorMessageCalled = false;
                        controller.OnAddVolunteer("TestSurname", "invalid-email");

                        // Assert - Verify error message was shown for invalid email
                        if (!mockUI.ShowErrorMessageCalled)
                        {
                            return false.Label("Error handling failed: ShowErrorMessage not called for invalid email");
                        }

                        // Test Case 4: Empty surname validation
                        mockUI.ShowErrorMessageCalled = false;
                        controller.OnAddVolunteer("", "test@example.com");

                        // Assert - Verify error message was shown for empty surname
                        if (!mockUI.ShowErrorMessageCalled)
                        {
                            return false.Label("Error handling failed: ShowErrorMessage not called for empty surname");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempAppDir))
                        {
                            try { Directory.Delete(tempAppDir, true); } catch { }
                        }
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Property 2: Preservation - UI Refresh and Send Emails Button State Updates Remain Unchanged
        /// 
        /// This test captures the ACTUAL UI update behavior on unfixed code.
        /// The fix should not change this behavior.
        /// 
        /// For any volunteer operation, the system SHALL:
        /// - Refresh the UI display when volunteers change
        /// - Update the send emails button state based on CanSendEmails()
        /// 
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void Property_UIUpdates_RemainsUnchanged()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                Arb.From(ValidSurnameGen()),
                Arb.From(ValidEmailGen()),
                (string newSurname, string newEmail) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create mock services
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Create controller (starts with empty volunteers)
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        // Reset mock UI call tracking
                        mockUI.DisplayVolunteerListCalled = false;
                        mockUI.EnableSendEmailsButtonCalled = false;

                        // Test Case 1: OnAddVolunteer updates UI
                        controller.OnAddVolunteer(newSurname, newEmail);

                        if (!mockUI.DisplayVolunteerListCalled)
                        {
                            return false.Label("UI update failed: DisplayVolunteerList not called after OnAddVolunteer");
                        }

                        if (!mockUI.EnableSendEmailsButtonCalled)
                        {
                            return false.Label("UI update failed: EnableSendEmailsButton not called after OnAddVolunteer");
                        }

                        // Test Case 2: OnVolunteerFileSelected updates UI
                        var externalFilePath = Path.Combine(Path.GetTempPath(), $"external_{Guid.NewGuid()}.json");
                        mockVolunteerManager.SaveVolunteers(externalFilePath, new Dictionary<string, string> { { "TestVolunteer", "test@example.com" } });

                        mockUI.DisplayVolunteerListCalled = false;
                        mockUI.EnableSendEmailsButtonCalled = false;

                        controller.OnVolunteerFileSelected(externalFilePath);

                        if (!mockUI.DisplayVolunteerListCalled)
                        {
                            File.Delete(externalFilePath);
                            return false.Label("UI update failed: DisplayVolunteerList not called after OnVolunteerFileSelected");
                        }

                        if (!mockUI.EnableSendEmailsButtonCalled)
                        {
                            File.Delete(externalFilePath);
                            return false.Label("UI update failed: EnableSendEmailsButton not called after OnVolunteerFileSelected");
                        }

                        File.Delete(externalFilePath);

                        // Test Case 3: OnGmailCredentialsUpdated updates send button state
                        mockUI.EnableSendEmailsButtonCalled = false;

                        controller.OnGmailCredentialsUpdated("test@gmail.com", "test-app-password");

                        if (!mockUI.EnableSendEmailsButtonCalled)
                        {
                            return false.Label("UI update failed: EnableSendEmailsButton not called after OnGmailCredentialsUpdated");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempAppDir))
                        {
                            try { Directory.Delete(tempAppDir, true); } catch { }
                        }
                    }
                }
            ).Check(config);
        }

        #region Mock Implementations

        private class MockEmailService : IEmailService
        {
            public Task<bool> SendVolunteerNotificationAsync(
                string toEmail,
                string volunteerSurname,
                List<Dictionary<string, string>> assignedRows,
                GmailCredentials credentials)
            {
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

        private class MockExcelManager : IExcelManager
        {
            private List<VolunteerAssignment> _assignments;

            public MockExcelManager(List<VolunteerAssignment> assignments)
            {
                _assignments = assignments;
            }

            public ExcelWorkbook OpenWorkbook(string filePath)
            {
                var package = new OfficeOpenXml.ExcelPackage();
                return new ExcelWorkbook(package);
            }

            public List<string> GetSheetNames(ExcelWorkbook workbook)
            {
                return new List<string> { "Sheet1" };
            }

            public Sheet GetSheetByName(ExcelWorkbook workbook, string sheetName)
            {
                var package = new OfficeOpenXml.ExcelPackage();
                var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                return new Sheet(worksheet);
            }

            public List<VolunteerAssignment> IdentifyVolunteerAssignments(Sheet sheet, Dictionary<string, string> volunteers)
            {
                return _assignments;
            }

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
            public void AppendLaboratoriData(Sheet targetSheet, Sheet laboratoriSheet, int startRow) { }
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

        private class MockVolunteerUI : IVolunteerUI
        {
            public bool DisplayVolunteerListCalled { get; set; }
            public bool EnableSendEmailsButtonCalled { get; set; }
            public bool ShowErrorMessageCalled { get; set; }

            public void DisplayVolunteerList(Dictionary<string, string> volunteers)
            {
                DisplayVolunteerListCalled = true;
            }

            public void DisplayGmailCredentials(string email, string password) { }

            public void DisplaySheetNames(List<string> sheetNames) { }

            public void EnableSendEmailsButton(bool enabled)
            {
                EnableSendEmailsButtonCalled = true;
            }

            public void ShowEmailProgress(string message) { }

            public void ShowEmailSummary(int successCount, int failureCount) { }

            public bool ConfirmAction(string message) => true;

            public void ShowErrorMessage(string message)
            {
                ShowErrorMessageCalled = true;
            }
        }

        #endregion
    }
}
