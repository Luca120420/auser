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
    /// Bug condition exploration tests for volunteer deletion persistence bug.
    /// These tests are EXPECTED TO FAIL on unfixed code - failure confirms the bug exists.
    /// **Validates: Requirements 2.1, 2.2, 2.3**
    /// </summary>
    [TestFixture]
    public class VolunteerDeletionPersistenceBugTests
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
        /// Property 1: Fault Condition - Deletions Persist to Internal Storage
        /// 
        /// CRITICAL: This test MUST FAIL on unfixed code - failure confirms the bug exists.
        /// DO NOT attempt to fix the test or the code when it fails.
        /// 
        /// This test encodes the expected behavior - it will validate the fix when it passes after implementation.
        /// 
        /// For any deletion operation (OnDeleteVolunteer or OnDeleteAllVolunteers),
        /// the system SHALL persist changes to internal storage (volunteers.json),
        /// ensuring deleted volunteers do not reappear after application restart.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3**
        /// </summary>
        [Test]
        public void Property_DeletionsPersistToInternalStorage()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> initialVolunteers) =>
                {
                    // Skip if we have fewer than 2 volunteers (need at least 2 to test individual deletion)
                    if (initialVolunteers.Count < 2)
                    {
                        return true.ToProperty().Label("Skipped: need at least 2 volunteers");
                    }

                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var dataFolder = Path.Combine(tempAppDir, "data");
                    var volunteersPath = Path.Combine(dataFolder, "volunteers.json");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);
                        Directory.CreateDirectory(dataFolder);

                        // Create mock services with real file I/O
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Save initial volunteers to internal storage
                        mockVolunteerManager.SaveVolunteers(volunteersPath, initialVolunteers);

                        // Create controller - this should load volunteers from internal storage
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        // Verify initial volunteers are loaded
                        var loadedVolunteers = controller.GetVolunteers();
                        if (loadedVolunteers.Count != initialVolunteers.Count)
                        {
                            return false.Label($"Initial load failed: expected {initialVolunteers.Count} volunteers, got {loadedVolunteers.Count}");
                        }

                        // Test Case 1: OnDeleteVolunteer removes volunteer from internal storage
                        var volunteerToDelete = initialVolunteers.Keys.First();
                        controller.OnDeleteVolunteer(volunteerToDelete);

                        // Verify internal storage is updated (volunteers.json should not contain deleted volunteer)
                        var volunteersAfterDelete = mockVolunteerManager.LoadVolunteers(volunteersPath);
                        if (volunteersAfterDelete.ContainsKey(volunteerToDelete))
                        {
                            return false.Label($"BUG DETECTED: OnDeleteVolunteer did not persist deletion to internal storage. Volunteer '{volunteerToDelete}' still exists in volunteers.json");
                        }

                        if (volunteersAfterDelete.Count != initialVolunteers.Count - 1)
                        {
                            return false.Label($"BUG DETECTED: Internal storage has incorrect count after deletion. Expected {initialVolunteers.Count - 1}, got {volunteersAfterDelete.Count}");
                        }

                        // Test Case 2: Deleted volunteer doesn't reappear after controller reinitialization (simulating restart)
                        var controller2 = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        var volunteersAfterRestart = controller2.GetVolunteers();
                        if (volunteersAfterRestart.ContainsKey(volunteerToDelete))
                        {
                            return false.Label($"BUG DETECTED: Deleted volunteer '{volunteerToDelete}' reappeared after restart simulation");
                        }

                        // Test Case 3: OnDeleteAllVolunteers clears internal storage
                        controller2.OnDeleteAllVolunteers();

                        // Verify internal storage is cleared (volunteers.json should be empty)
                        var volunteersAfterDeleteAll = mockVolunteerManager.LoadVolunteers(volunteersPath);
                        if (volunteersAfterDeleteAll.Count != 0)
                        {
                            return false.Label($"BUG DETECTED: OnDeleteAllVolunteers did not clear internal storage. volunteers.json still contains {volunteersAfterDeleteAll.Count} volunteers");
                        }

                        // Test Case 4: No volunteers reappear after restart following delete all
                        var controller3 = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        var volunteersAfterFinalRestart = controller3.GetVolunteers();
                        if (volunteersAfterFinalRestart.Count != 0)
                        {
                            return false.Label($"BUG DETECTED: {volunteersAfterFinalRestart.Count} volunteers reappeared after restart following delete all");
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
        /// Property 1 (Extended): Deletions Persist to Internal Storage After Import
        /// 
        /// CRITICAL: This test MUST FAIL on unfixed code - failure confirms the bug exists.
        /// 
        /// When volunteers are imported from an external file and then deleted,
        /// the deletions should persist to internal storage (volunteers.json).
        /// The external file is NOT updated because it's only used for import.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3**
        /// </summary>
        [Test]
        public void Property_DeletionsPersistToInternalStorageAfterImport()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 50;

            Prop.ForAll(
                Arb.From(VolunteerDictionaryGen()),
                (Dictionary<string, string> initialVolunteers) =>
                {
                    // Skip if we have fewer than 2 volunteers
                    if (initialVolunteers.Count < 2)
                    {
                        return true.ToProperty().Label("Skipped: need at least 2 volunteers");
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
                        Directory.CreateDirectory(dataFolder);

                        // Create external volunteer file for import
                        var mockVolunteerManager = new MockVolunteerManagerWithRealIO(dataFolder);
                        mockVolunteerManager.SaveVolunteers(externalFilePath, initialVolunteers);

                        var mockEmailService = new MockEmailService();
                        var mockConfigService = new MockConfigurationServiceWithAppDir(tempAppDir);
                        var mockExcelManager = new MockExcelManager(new List<VolunteerAssignment>());
                        var mockUI = new MockVolunteerUI();

                        // Create controller and import from external file
                        var controller = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder); // Pass custom data folder for test isolation

                        controller.OnVolunteerFileSelected(externalFilePath);

                        // Act - Delete a volunteer
                        var volunteerToDelete = initialVolunteers.Keys.First();
                        controller.OnDeleteVolunteer(volunteerToDelete);

                        // Assert - Verify internal storage is updated
                        var internalVolunteers = mockVolunteerManager.LoadVolunteers(volunteersPath);
                        if (internalVolunteers.ContainsKey(volunteerToDelete))
                        {
                            return false.Label($"BUG DETECTED: Internal storage was not updated: volunteer '{volunteerToDelete}' still exists in volunteers.json");
                        }

                        if (internalVolunteers.Count != initialVolunteers.Count - 1)
                        {
                            return false.Label($"BUG DETECTED: Internal storage has incorrect count. Expected {initialVolunteers.Count - 1}, got {internalVolunteers.Count}");
                        }

                        // Verify deleted volunteer doesn't reappear after restart
                        var controller2 = new VolunteerNotificationController(
                            mockVolunteerManager,
                            mockEmailService,
                            mockConfigService,
                            mockExcelManager,
                            mockUI,
                            dataFolder);

                        var volunteersAfterRestart = controller2.GetVolunteers();
                        if (volunteersAfterRestart.ContainsKey(volunteerToDelete))
                        {
                            return false.Label($"BUG DETECTED: Deleted volunteer '{volunteerToDelete}' reappeared after restart");
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

        private class MockVolunteerUI : IVolunteerUI
        {
            public void DisplayVolunteerList(Dictionary<string, string> volunteers) { }
            public void DisplayGmailCredentials(string email, string password) { }
            public void DisplaySheetNames(List<string> sheetNames) { }
            public void EnableSendEmailsButton(bool enabled) { }
            public void ShowEmailProgress(string message) { }
            public void ShowEmailSummary(int successCount, int failureCount) { }
            public bool ConfirmAction(string message) => true;
            public void ShowErrorMessage(string message) { }
        }

        #endregion
    }
}
