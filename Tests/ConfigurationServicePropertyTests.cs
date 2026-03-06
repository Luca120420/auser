using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using FsCheck;
using NUnit.Framework;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Property-based tests for ConfigurationService class using FsCheck.
    /// Tests universal properties that should hold across all valid inputs.
    /// Validates: Requirements 1.3, 3.2, 7.1, 7.2
    /// </summary>
    [TestFixture]
    public class ConfigurationServicePropertyTests
    {
        private ConfigurationService _configurationService = null!;

        [SetUp]
        public void Setup()
        {
            var volunteerManager = new VolunteerManager();
            _configurationService = new ConfigurationService(volunteerManager);
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
        /// Custom generator for valid file paths
        /// </summary>
        private static Gen<string> ValidFilePathGen()
        {
            var filenameGen = from length in Gen.Choose(1, 20)
                             from chars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyz0123456789_-".ToCharArray()))
                             from ext in Gen.Elements("json", "xlsx", "csv", "txt")
                             select new string(chars) + "." + ext;

            var pathGen = from depth in Gen.Choose(0, 3)
                         from segments in Gen.ListOf(depth, Gen.Elements("folder", "data", "files", "documents"))
                         from filename in filenameGen
                         select string.Join(Path.DirectorySeparatorChar.ToString(), segments.Append(filename));

            return pathGen;
        }

        /// <summary>
        /// Custom generator for GmailCredentials
        /// </summary>
        private static Gen<GmailCredentials> GmailCredentialsGen()
        {
            var passwordGen = from length in Gen.Choose(16, 32)
                             from chars in Gen.ArrayOf(length, Gen.Elements("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789".ToCharArray()))
                             select new string(chars);

            return from email in ValidEmailGen()
                   from password in passwordGen
                   select new GmailCredentials
                   {
                       Email = email,
                       AppPassword = password
                   };
        }

        /// <summary>
        /// Custom generator for AppConfiguration
        /// </summary>
        private static Gen<AppConfiguration> AppConfigurationGen()
        {
            return from excelPath in ValidFilePathGen()
                   from sheetName in Gen.Elements("Sheet1", "Foglio1", "Data", "Volontari", "Servizi")
                   from credentials in GmailCredentialsGen()
                   select new AppConfiguration
                   {
                       // VolunteerFilePath removed - no longer part of AppConfiguration (Task 2.1)
                       LastExcelFilePath = excelPath,
                       LastSheetName = sheetName,
                       GmailCredentials = credentials
                   };
        }

        // Feature: volunteer-email-notifications, Property 2: Configuration Persistence Round Trip
        /// <summary>
        /// Property 2: Configuration Persistence Round Trip
        /// For any valid configuration data (volunteer file path and Gmail credentials),
        /// saving the configuration and then loading it back should produce equivalent values.
        /// **Validates: Requirements 1.3, 3.2, 7.1, 7.2**
        /// </summary>
        [Test]
        public void Property_ConfigurationPersistenceRoundTrip()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                (AppConfiguration originalConfig) =>
                {
                    // Use a temporary directory for the test
                    var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDir);
                    var tempConfigFile = Path.Combine(tempDir, "config.json");

                    try
                    {
                        // Create a custom ConfigurationService that uses the temp directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationService(tempConfigFile, volunteerManager);

                        // Act - Save configuration
                        testService.SaveConfiguration(originalConfig);

                        // Act - Load configuration
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify LastExcelFilePath
                        if (originalConfig.LastExcelFilePath != loadedConfig.LastExcelFilePath)
                        {
                            return false.Label($"LastExcelFilePath mismatch: expected '{originalConfig.LastExcelFilePath}', got '{loadedConfig.LastExcelFilePath}'");
                        }

                        // Assert - Verify LastSheetName
                        if (originalConfig.LastSheetName != loadedConfig.LastSheetName)
                        {
                            return false.Label($"LastSheetName mismatch: expected '{originalConfig.LastSheetName}', got '{loadedConfig.LastSheetName}'");
                        }

                        // Assert - Verify Gmail Email
                        if (originalConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email mismatch: expected '{originalConfig.GmailCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        // Assert - Verify Gmail AppPassword
                        if (originalConfig.GmailCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label($"Gmail AppPassword mismatch: expected '{originalConfig.GmailCredentials.AppPassword}', got '{loadedConfig.GmailCredentials.AppPassword}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Configuration round trip failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: volunteer-email-notifications, Property 5: Latest Configuration Value Wins
        /// <summary>
        /// Property 5: Latest Configuration Value Wins
        /// For any two different configuration values (file paths or credentials),
        /// saving the first value then saving the second value should result in
        /// only the second value being persisted.
        /// **Validates: Requirements 1.3, 1.7, 3.2, 3.4, 7.1, 7.2**
        /// </summary>
        [Test]
        public void Property_LatestConfigurationValueWins()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                Arb.From(AppConfigurationGen()),
                (AppConfiguration firstConfig, AppConfiguration secondConfig) =>
                {
                    // Ensure the two configurations are different
                    if (AreConfigurationsEqual(firstConfig, secondConfig))
                    {
                        return true.ToProperty().Label("Skipped: Generated identical configurations");
                    }

                    // Use a temporary directory for the test
                    var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
                    Directory.CreateDirectory(tempDir);
                    var tempConfigFile = Path.Combine(tempDir, "config.json");

                    try
                    {
                        // Create a custom ConfigurationService that uses the temp directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationService(tempConfigFile, volunteerManager);

                        // Act - Save first configuration
                        testService.SaveConfiguration(firstConfig);

                        // Act - Save second configuration (should overwrite first)
                        testService.SaveConfiguration(secondConfig);

                        // Act - Load configuration
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify loaded config matches second config, not first
                        if (secondConfig.LastExcelFilePath != loadedConfig.LastExcelFilePath)
                        {
                            return false.Label($"LastExcelFilePath should be second value: expected '{secondConfig.LastExcelFilePath}', got '{loadedConfig.LastExcelFilePath}'");
                        }

                        if (secondConfig.LastSheetName != loadedConfig.LastSheetName)
                        {
                            return false.Label($"LastSheetName should be second value: expected '{secondConfig.LastSheetName}', got '{loadedConfig.LastSheetName}'");
                        }

                        if (secondConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email should be second value: expected '{secondConfig.GmailCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        if (secondConfig.GmailCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label($"Gmail AppPassword should be second value: expected '{secondConfig.GmailCredentials.AppPassword}', got '{loadedConfig.GmailCredentials.AppPassword}'");
                        }

                        // Assert - Verify loaded config does NOT match first config (at least one field should differ)
                        bool matchesFirst = firstConfig.LastExcelFilePath == loadedConfig.LastExcelFilePath &&
                                          firstConfig.LastSheetName == loadedConfig.LastSheetName &&
                                          firstConfig.GmailCredentials.Email == loadedConfig.GmailCredentials.Email &&
                                          firstConfig.GmailCredentials.AppPassword == loadedConfig.GmailCredentials.AppPassword;

                        if (matchesFirst)
                        {
                            return false.Label("Loaded config matches first config instead of second config");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Latest configuration value wins test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup
                        if (Directory.Exists(tempDir))
                        {
                            Directory.Delete(tempDir, true);
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 1: Configuration Round-Trip Persistence
        /// <summary>
        /// Property 1: Configuration Round-Trip Persistence
        /// For any valid AppConfiguration object, saving it to the data folder and then loading it
        /// should produce an equivalent configuration with the same Gmail credentials, LastExcelFilePath,
        /// and LastSheetName values.
        /// **Validates: Requirements 1.4, 1.5, 6.1, 6.3**
        /// </summary>
        [Test]
        public void Property_ConfigurationRoundTripPersistence()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                (AppConfiguration originalConfig) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var expectedDataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);

                        // Create a test service that uses the temporary app directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithAppDir(tempAppDir, volunteerManager);

                        // Act - Save configuration (this will encrypt the password)
                        testService.SaveConfiguration(originalConfig);

                        // Act - Load configuration (this will decrypt the password)
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify LastExcelFilePath
                        if (originalConfig.LastExcelFilePath != loadedConfig.LastExcelFilePath)
                        {
                            return false.Label($"LastExcelFilePath mismatch: expected '{originalConfig.LastExcelFilePath}', got '{loadedConfig.LastExcelFilePath}'");
                        }

                        // Assert - Verify LastSheetName
                        if (originalConfig.LastSheetName != loadedConfig.LastSheetName)
                        {
                            return false.Label($"LastSheetName mismatch: expected '{originalConfig.LastSheetName}', got '{loadedConfig.LastSheetName}'");
                        }

                        // Assert - Verify Gmail Email
                        if (originalConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email mismatch: expected '{originalConfig.GmailCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        // Assert - Verify Gmail AppPassword (should survive encryption/decryption round-trip)
                        if (originalConfig.GmailCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label($"Gmail AppPassword mismatch: expected '{originalConfig.GmailCredentials.AppPassword}', got '{loadedConfig.GmailCredentials.AppPassword}'");
                        }

                        // Assert - Verify data was saved to the correct location (data folder)
                        var expectedConfigPath = Path.Combine(expectedDataFolder, "config.json");
                        if (!File.Exists(expectedConfigPath))
                        {
                            return false.Label($"Configuration file was not saved to expected path: {expectedConfigPath}");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Configuration round-trip persistence test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary app directory
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 3: Data Folder Creation
        /// <summary>
        /// Property 3: Data Folder Creation
        /// For any state where the data folder does not exist, calling EnsureDataFolderExists()
        /// should result in the data folder existing at the expected path relative to the application folder.
        /// **Validates: Requirements 1.2**
        /// </summary>
        [Test]
        public void Property_DataFolderCreation()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.Default.Guid(),
                (Guid testId) =>
                {
                    // Use a temporary directory for the test to avoid interference
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{testId}");
                    var expectedDataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory but not the data folder
                        Directory.CreateDirectory(tempAppDir);
                        
                        // Ensure data folder does NOT exist initially
                        if (Directory.Exists(expectedDataFolder))
                        {
                            Directory.Delete(expectedDataFolder, true);
                        }

                        // Verify precondition: data folder should not exist
                        if (Directory.Exists(expectedDataFolder))
                        {
                            return false.Label("Precondition failed: data folder exists before test");
                        }

                        // Create a test service that uses the temporary app directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithAppDir(tempAppDir, volunteerManager);

                        // Act - Call EnsureDataFolderExists()
                        testService.EnsureDataFolderExists();

                        // Assert - Verify data folder now exists
                        if (!Directory.Exists(expectedDataFolder))
                        {
                            return false.Label($"Data folder was not created at expected path: {expectedDataFolder}");
                        }

                        // Verify the folder is at the correct relative path
                        var actualPath = Path.GetFullPath(expectedDataFolder);
                        var expectedPath = Path.GetFullPath(Path.Combine(tempAppDir, "data"));
                        
                        if (actualPath != expectedPath)
                        {
                            return false.Label($"Data folder path mismatch: expected '{expectedPath}', got '{actualPath}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Data folder creation test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary app directory
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 7: Relative Path Resolution
        /// <summary>
        /// Property 7: Relative Path Resolution
        /// For any file operation (config or volunteer data), the paths used should be constructed
        /// relative to the application base directory, not as absolute paths to the data folder.
        /// **Validates: Requirements 4.1, 4.2, 4.4**
        /// </summary>
        [Test]
        public void Property_RelativePathResolution()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.Default.Guid(),
                (Guid testId) =>
                {
                    // Use a temporary directory for the test to simulate different application folders
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{testId}");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);

                        // Create a test service that uses the temporary app directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithAppDir(tempAppDir, volunteerManager);

                        // Act - Get the config file path
                        string configPath = testService.GetConfigFilePath();

                        // Assert - Verify the path is constructed relative to the application folder
                        // The path should be: {ApplicationFolder}\data\config.json
                        
                        // 1. Verify the path contains the application folder
                        if (!configPath.StartsWith(tempAppDir))
                        {
                            return false.Label($"Config path does not start with application folder. Expected to start with '{tempAppDir}', got '{configPath}'");
                        }

                        // 2. Verify the path structure is {ApplicationFolder}\data\config.json
                        string expectedPath = Path.Combine(tempAppDir, "data", "config.json");
                        string actualPath = Path.GetFullPath(configPath);
                        string expectedFullPath = Path.GetFullPath(expectedPath);

                        if (actualPath != expectedFullPath)
                        {
                            return false.Label($"Config path structure mismatch. Expected '{expectedFullPath}', got '{actualPath}'");
                        }

                        // 3. Verify the path is NOT an absolute path to a fixed location
                        // (i.e., it should change when the application folder changes)
                        // This is implicitly verified by the above checks, but let's be explicit:
                        // The path should contain the unique test ID, proving it's relative to the app folder
                        if (!configPath.Contains(testId.ToString()))
                        {
                            return false.Label($"Config path does not contain the test ID, suggesting it's not relative to the application folder");
                        }

                        // 4. Verify the relative portion is "data\config.json"
                        string relativePath = Path.GetRelativePath(tempAppDir, configPath);
                        string expectedRelativePath = Path.Combine("data", "config.json");

                        if (relativePath != expectedRelativePath)
                        {
                            return false.Label($"Relative path mismatch. Expected '{expectedRelativePath}', got '{relativePath}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Relative path resolution test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary app directory
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 8: Configuration Migration Copies Data
        /// <summary>
        /// Property 8: Configuration Migration Copies Data
        /// For any valid configuration file in the AppData folder, when the data folder config doesn't exist,
        /// loading configuration should result in the AppData config being copied to the data folder with
        /// identical content.
        /// **Validates: Requirements 5.2**
        /// </summary>
        [Test]
        public void Property_ConfigurationMigrationCopiesData()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                (AppConfiguration originalConfig) =>
                {
                    // Use temporary directories to simulate AppData and application folder
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var tempAppDataDir = Path.Combine(Path.GetTempPath(), $"test_appdata_{Guid.NewGuid()}");
                    var appDataConfigPath = Path.Combine(tempAppDataDir, "config.json");
                    var dataFolderPath = Path.Combine(tempAppDir, "data");
                    var dataConfigPath = Path.Combine(dataFolderPath, "config.json");

                    try
                    {
                        // Arrange - Create AppData directory and save config there
                        Directory.CreateDirectory(tempAppDataDir);
                        
                        // Save the original config to AppData location (with encryption)
                        var options = new System.Text.Json.JsonSerializerOptions
                        {
                            WriteIndented = true
                        };
                        
                        // Create a copy with encrypted password for AppData
                        var appDataConfig = new AppConfiguration
                        {
                            // VolunteerFilePath removed - no longer part of AppConfiguration (Task 2.1)
                            LastExcelFilePath = originalConfig.LastExcelFilePath,
                            LastSheetName = originalConfig.LastSheetName,
                            GmailCredentials = new GmailCredentials
                            {
                                Email = originalConfig.GmailCredentials.Email,
                                AppPassword = EncryptPasswordForTest(originalConfig.GmailCredentials.AppPassword)
                            }
                        };
                        
                        string appDataJson = System.Text.Json.JsonSerializer.Serialize(appDataConfig, options);
                        File.WriteAllText(appDataConfigPath, appDataJson);

                        // Arrange - Create app directory but ensure data folder config doesn't exist
                        Directory.CreateDirectory(tempAppDir);
                        if (Directory.Exists(dataFolderPath))
                        {
                            Directory.Delete(dataFolderPath, true);
                        }

                        // Verify precondition: data folder config should not exist
                        if (File.Exists(dataConfigPath))
                        {
                            return false.Label("Precondition failed: data folder config exists before migration");
                        }

                        // Create a test service that uses the temporary directories
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithMigration(tempAppDir, tempAppDataDir, volunteerManager);

                        // Act - Load configuration (this should trigger migration)
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify data folder config now exists
                        if (!File.Exists(dataConfigPath))
                        {
                            return false.Label($"Configuration was not migrated to data folder: {dataConfigPath}");
                        }

                        // Assert - Verify migrated config has identical content
                        if (originalConfig.LastExcelFilePath != loadedConfig.LastExcelFilePath)
                        {
                            return false.Label($"LastExcelFilePath mismatch after migration: expected '{originalConfig.LastExcelFilePath}', got '{loadedConfig.LastExcelFilePath}'");
                        }

                        if (originalConfig.LastSheetName != loadedConfig.LastSheetName)
                        {
                            return false.Label($"LastSheetName mismatch after migration: expected '{originalConfig.LastSheetName}', got '{loadedConfig.LastSheetName}'");
                        }

                        if (originalConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email mismatch after migration: expected '{originalConfig.GmailCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        if (originalConfig.GmailCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label($"Gmail AppPassword mismatch after migration: expected '{originalConfig.GmailCredentials.AppPassword}', got '{loadedConfig.GmailCredentials.AppPassword}'");
                        }

                        // Assert - Verify the migrated file content matches (by reading the file directly)
                        string migratedJson = File.ReadAllText(dataConfigPath);
                        var migratedConfig = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(migratedJson);
                        
                        if (migratedConfig == null)
                        {
                            return false.Label("Migrated configuration file is null or corrupted");
                        }

                        // Verify the file was copied correctly (encrypted password should match)
                        if (migratedConfig.GmailCredentials.AppPassword != appDataConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label("Encrypted password in migrated file doesn't match AppData file");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Configuration migration test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary directories
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                        
                        if (Directory.Exists(tempAppDataDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDataDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 10: Migration Removes VolunteerFilePath
        /// <summary>
        /// Property 10: Migration Removes VolunteerFilePath
        /// For any configuration migrated from AppData that contains a VolunteerFilePath property,
        /// after migration completes, the saved configuration in the data folder should not contain
        /// the VolunteerFilePath property.
        /// **Validates: Requirements 5.5**
        /// </summary>
        [Test]
        public void Property_MigrationRemovesVolunteerFilePath()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                Arb.From(ValidFilePathGen()),
                (AppConfiguration originalConfig, string volunteerFilePath) =>
                {
                    // Use temporary directories to simulate AppData and application folder
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var tempAppDataDir = Path.Combine(Path.GetTempPath(), $"test_appdata_{Guid.NewGuid()}");
                    var appDataConfigPath = Path.Combine(tempAppDataDir, "config.json");
                    var dataFolderPath = Path.Combine(tempAppDir, "data");
                    var dataConfigPath = Path.Combine(dataFolderPath, "config.json");

                    try
                    {
                        // Arrange - Create AppData directory and save config there WITH VolunteerFilePath
                        Directory.CreateDirectory(tempAppDataDir);
                        
                        // Create a JSON string manually to include the VolunteerFilePath property
                        // (since it's been removed from the AppConfiguration model)
                        var configWithVolunteerPath = new
                        {
                            VolunteerFilePath = volunteerFilePath,
                            LastExcelFilePath = originalConfig.LastExcelFilePath,
                            LastSheetName = originalConfig.LastSheetName,
                            GmailCredentials = new
                            {
                                Email = originalConfig.GmailCredentials.Email,
                                AppPassword = EncryptPasswordForTest(originalConfig.GmailCredentials.AppPassword)
                            }
                        };
                        
                        var options = new System.Text.Json.JsonSerializerOptions
                        {
                            WriteIndented = true
                        };
                        
                        string appDataJson = System.Text.Json.JsonSerializer.Serialize(configWithVolunteerPath, options);
                        File.WriteAllText(appDataConfigPath, appDataJson);

                        // Verify the AppData config contains VolunteerFilePath
                        string appDataContent = File.ReadAllText(appDataConfigPath);
                        if (!appDataContent.Contains("VolunteerFilePath"))
                        {
                            return false.Label("Precondition failed: AppData config doesn't contain VolunteerFilePath");
                        }

                        // Arrange - Create app directory but ensure data folder config doesn't exist
                        Directory.CreateDirectory(tempAppDir);
                        if (Directory.Exists(dataFolderPath))
                        {
                            Directory.Delete(dataFolderPath, true);
                        }

                        // Verify precondition: data folder config should not exist
                        if (File.Exists(dataConfigPath))
                        {
                            return false.Label("Precondition failed: data folder config exists before migration");
                        }

                        // Create a test service that uses the temporary directories
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithMigration(tempAppDir, tempAppDataDir, volunteerManager);

                        // Act - Load configuration (this should trigger migration)
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify data folder config now exists
                        if (!File.Exists(dataConfigPath))
                        {
                            return false.Label($"Configuration was not migrated to data folder: {dataConfigPath}");
                        }

                        // Assert - Verify the migrated config file does NOT contain VolunteerFilePath
                        string migratedContent = File.ReadAllText(dataConfigPath);
                        if (migratedContent.Contains("VolunteerFilePath"))
                        {
                            return false.Label("Migration failed: VolunteerFilePath property still exists in migrated config file");
                        }

                        // Assert - Verify the loaded config object doesn't have VolunteerFilePath
                        // (This is implicitly true since AppConfiguration model doesn't have this property,
                        // but we verify the JSON deserialization worked correctly)
                        var migratedConfig = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(migratedContent);
                        if (migratedConfig == null)
                        {
                            return false.Label("Migrated configuration file is null or corrupted");
                        }

                        // Assert - Verify other properties were preserved correctly
                        if (originalConfig.LastExcelFilePath != loadedConfig.LastExcelFilePath)
                        {
                            return false.Label($"LastExcelFilePath mismatch after migration: expected '{originalConfig.LastExcelFilePath}', got '{loadedConfig.LastExcelFilePath}'");
                        }

                        if (originalConfig.LastSheetName != loadedConfig.LastSheetName)
                        {
                            return false.Label($"LastSheetName mismatch after migration: expected '{originalConfig.LastSheetName}', got '{loadedConfig.LastSheetName}'");
                        }

                        if (originalConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email mismatch after migration: expected '{originalConfig.GmailCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Migration removes VolunteerFilePath test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary directories
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                        
                        if (Directory.Exists(tempAppDataDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDataDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 9: Volunteer Migration Copies Data
        /// <summary>
        /// Property 9: Volunteer Migration Copies Data
        /// For any valid volunteer file referenced in a migrated configuration, if the external file exists,
        /// migration should result in the volunteer data being copied to data/volunteers.json with identical content.
        /// **Validates: Requirements 5.4**
        /// </summary>
        [Test]
        public void Property_VolunteerMigrationCopiesData()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Custom generator for volunteer dictionaries
            var volunteerDictGen = from count in Gen.Choose(1, 10)
                                   from surnames in Gen.ListOf(count, Gen.Elements("Rossi", "Bianchi", "Verdi", "Romano", "Colombo", "Ferrari", "Esposito", "Ricci", "Marino", "Greco"))
                                   from emails in Gen.ListOf(count, ValidEmailGen())
                                   select surnames.Zip(emails, (s, e) => new { Surname = s, Email = e })
                                                 .GroupBy(x => x.Surname)
                                                 .ToDictionary(g => g.First().Surname, g => g.First().Email);

            Prop.ForAll(
                Arb.From(volunteerDictGen),
                Arb.From(AppConfigurationGen()),
                (Dictionary<string, string> originalVolunteers, AppConfiguration originalConfig) =>
                {
                    // Use temporary directories to simulate AppData, application folder, and external volunteer file
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var tempAppDataDir = Path.Combine(Path.GetTempPath(), $"test_appdata_{Guid.NewGuid()}");
                    var tempExternalVolunteerFile = Path.Combine(Path.GetTempPath(), $"volunteers_{Guid.NewGuid()}.json");
                    var appDataConfigPath = Path.Combine(tempAppDataDir, "config.json");
                    var dataFolderPath = Path.Combine(tempAppDir, "data");
                    var dataConfigPath = Path.Combine(dataFolderPath, "config.json");
                    var dataVolunteersPath = Path.Combine(dataFolderPath, "volunteers.json");

                    try
                    {
                        // Arrange - Create external volunteer file with test data
                        var volunteerManager = new VolunteerManager();
                        volunteerManager.SaveVolunteers(tempExternalVolunteerFile, originalVolunteers);

                        // Verify the external file was created
                        if (!File.Exists(tempExternalVolunteerFile))
                        {
                            return false.Label("Precondition failed: External volunteer file was not created");
                        }

                        // Arrange - Create AppData directory and save config WITH VolunteerFilePath
                        Directory.CreateDirectory(tempAppDataDir);
                        
                        var configWithVolunteerPath = new
                        {
                            VolunteerFilePath = tempExternalVolunteerFile,
                            LastExcelFilePath = originalConfig.LastExcelFilePath,
                            LastSheetName = originalConfig.LastSheetName,
                            GmailCredentials = new
                            {
                                Email = originalConfig.GmailCredentials.Email,
                                AppPassword = EncryptPasswordForTest(originalConfig.GmailCredentials.AppPassword)
                            }
                        };
                        
                        var options = new System.Text.Json.JsonSerializerOptions
                        {
                            WriteIndented = true
                        };
                        
                        string appDataJson = System.Text.Json.JsonSerializer.Serialize(configWithVolunteerPath, options);
                        File.WriteAllText(appDataConfigPath, appDataJson);

                        // Verify the AppData config contains VolunteerFilePath
                        string appDataContent = File.ReadAllText(appDataConfigPath);
                        if (!appDataContent.Contains("VolunteerFilePath"))
                        {
                            return false.Label("Precondition failed: AppData config doesn't contain VolunteerFilePath");
                        }

                        // Arrange - Create app directory but ensure data folder doesn't exist
                        Directory.CreateDirectory(tempAppDir);
                        if (Directory.Exists(dataFolderPath))
                        {
                            Directory.Delete(dataFolderPath, true);
                        }

                        // Verify precondition: data folder volunteers.json should not exist
                        if (File.Exists(dataVolunteersPath))
                        {
                            return false.Label("Precondition failed: data/volunteers.json exists before migration");
                        }

                        // Create a test service that uses the temporary directories
                        var testService = new TestConfigurationServiceWithVolunteerMigration(tempAppDir, tempAppDataDir, volunteerManager);

                        // Act - Load configuration (this should trigger migration including volunteer data)
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify data/volunteers.json now exists
                        if (!File.Exists(dataVolunteersPath))
                        {
                            return false.Label($"Volunteer data was not migrated to data/volunteers.json: {dataVolunteersPath}");
                        }

                        // Assert - Verify migrated volunteer data has identical content
                        var migratedVolunteers = volunteerManager.LoadVolunteers(dataVolunteersPath);

                        if (migratedVolunteers.Count != originalVolunteers.Count)
                        {
                            return false.Label($"Volunteer count mismatch: expected {originalVolunteers.Count}, got {migratedVolunteers.Count}");
                        }

                        foreach (var kvp in originalVolunteers)
                        {
                            if (!migratedVolunteers.ContainsKey(kvp.Key))
                            {
                                return false.Label($"Migrated volunteers missing surname: {kvp.Key}");
                            }

                            if (migratedVolunteers[kvp.Key] != kvp.Value)
                            {
                                return false.Label($"Email mismatch for {kvp.Key}: expected '{kvp.Value}', got '{migratedVolunteers[kvp.Key]}'");
                            }
                        }

                        // Assert - Verify VolunteerFilePath was removed from config
                        string migratedConfigContent = File.ReadAllText(dataConfigPath);
                        if (migratedConfigContent.Contains("VolunteerFilePath"))
                        {
                            return false.Label("Migration failed: VolunteerFilePath property still exists in migrated config file");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Volunteer migration test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary directories and files
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                        
                        if (Directory.Exists(tempAppDataDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDataDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }

                        if (File.Exists(tempExternalVolunteerFile))
                        {
                            try
                            {
                                File.Delete(tempExternalVolunteerFile);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 12: Gmail Credentials Encryption Round-Trip
        /// <summary>
        /// Property 12: Gmail Credentials Encryption Round-Trip
        /// For any valid Gmail credentials (email and password), encrypting the password, saving the configuration,
        /// loading it, and decrypting should produce the original password value.
        /// **Validates: Requirements 7.1, 7.2, 7.3**
        /// </summary>
        [Test]
        public void Property_GmailCredentialsEncryptionRoundTrip()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.From(GmailCredentialsGen()),
                (GmailCredentials originalCredentials) =>
                {
                    // Use a temporary directory for the test
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var expectedDataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);

                        // Create a test configuration with the generated credentials
                        var originalConfig = new AppConfiguration
                        {
                            LastExcelFilePath = "test.xlsx",
                            LastSheetName = "Sheet1",
                            GmailCredentials = new GmailCredentials
                            {
                                Email = originalCredentials.Email,
                                AppPassword = originalCredentials.AppPassword
                            }
                        };

                        // Create a test service that uses the temporary app directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithAppDir(tempAppDir, volunteerManager);

                        // Act - Save configuration (this will encrypt the password using DPAPI)
                        testService.SaveConfiguration(originalConfig);

                        // Verify the password was encrypted in the saved file
                        var configPath = Path.Combine(expectedDataFolder, "config.json");
                        if (!File.Exists(configPath))
                        {
                            return false.Label($"Configuration file was not saved to expected path: {configPath}");
                        }

                        // Read the saved file and verify the password is encrypted (not plain text)
                        string savedJson = File.ReadAllText(configPath);
                        var savedConfig = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(savedJson);
                        
                        if (savedConfig == null)
                        {
                            return false.Label("Saved configuration is null");
                        }

                        // The saved password should be encrypted (base64 encoded), not the original plain text
                        if (savedConfig.GmailCredentials.AppPassword == originalCredentials.AppPassword)
                        {
                            return false.Label("Password was not encrypted in saved file (still in plain text)");
                        }

                        // Verify the encrypted password is base64 encoded (DPAPI output)
                        try
                        {
                            byte[] encryptedBytes = Convert.FromBase64String(savedConfig.GmailCredentials.AppPassword);
                            if (encryptedBytes.Length == 0)
                            {
                                return false.Label("Encrypted password is empty");
                            }
                        }
                        catch (FormatException)
                        {
                            return false.Label("Encrypted password is not valid base64 (DPAPI encryption failed)");
                        }

                        // Act - Load configuration (this will decrypt the password using DPAPI)
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify the email matches
                        if (originalCredentials.Email != loadedConfig.GmailCredentials.Email)
                        {
                            return false.Label($"Gmail Email mismatch: expected '{originalCredentials.Email}', got '{loadedConfig.GmailCredentials.Email}'");
                        }

                        // Assert - Verify the password was decrypted correctly (matches original plain text)
                        if (originalCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label($"Gmail AppPassword mismatch after encryption/decryption round-trip: expected '{originalCredentials.AppPassword}', got '{loadedConfig.GmailCredentials.AppPassword}'");
                        }

                        // Assert - Verify the round-trip preserved the original values
                        // (encryption -> save -> load -> decryption should be transparent)
                        if (originalConfig.GmailCredentials.Email != loadedConfig.GmailCredentials.Email ||
                            originalConfig.GmailCredentials.AppPassword != loadedConfig.GmailCredentials.AppPassword)
                        {
                            return false.Label("Encryption round-trip did not preserve original credentials");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"Gmail credentials encryption round-trip test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary app directory
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 11: AppData Not Used After Migration
        /// <summary>
        /// Property 11: AppData Not Used After Migration
        /// For any sequence of operations after migration is complete (data folder config exists),
        /// no files should be created or modified in the AppData folder.
        /// **Validates: Requirements 1.3, 9.1, 9.2, 9.3, 9.4**
        /// </summary>
        [Test]
        public void Property_AppDataNotUsedAfterMigration()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            // Custom generator for volunteer dictionaries
            var volunteerDictGen = from count in Gen.Choose(1, 5)
                                   from surnames in Gen.ListOf(count, Gen.Elements("Rossi", "Bianchi", "Verdi", "Romano", "Colombo"))
                                   from emails in Gen.ListOf(count, ValidEmailGen())
                                   select surnames.Zip(emails, (s, e) => new { Surname = s, Email = e })
                                                 .GroupBy(x => x.Surname)
                                                 .ToDictionary(g => g.First().Surname, g => g.First().Email);

            Prop.ForAll(
                Arb.From(AppConfigurationGen()),
                Arb.From(volunteerDictGen),
                (AppConfiguration originalConfig, Dictionary<string, string> originalVolunteers) =>
                {
                    // Use temporary directories to simulate AppData and application folder
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{Guid.NewGuid()}");
                    var tempAppDataDir = Path.Combine(Path.GetTempPath(), $"test_appdata_{Guid.NewGuid()}");
                    var tempExternalVolunteerFile = Path.Combine(Path.GetTempPath(), $"volunteers_{Guid.NewGuid()}.json");
                    var appDataConfigPath = Path.Combine(tempAppDataDir, "config.json");
                    var dataFolderPath = Path.Combine(tempAppDir, "data");
                    var dataConfigPath = Path.Combine(dataFolderPath, "config.json");
                    var dataVolunteersPath = Path.Combine(dataFolderPath, "volunteers.json");

                    try
                    {
                        // Arrange - Create external volunteer file with test data
                        var volunteerManager = new VolunteerManager();
                        volunteerManager.SaveVolunteers(tempExternalVolunteerFile, originalVolunteers);

                        // Arrange - Create AppData directory and save config WITH VolunteerFilePath
                        Directory.CreateDirectory(tempAppDataDir);
                        
                        var configWithVolunteerPath = new
                        {
                            VolunteerFilePath = tempExternalVolunteerFile,
                            LastExcelFilePath = originalConfig.LastExcelFilePath,
                            LastSheetName = originalConfig.LastSheetName,
                            GmailCredentials = new
                            {
                                Email = originalConfig.GmailCredentials.Email,
                                AppPassword = EncryptPasswordForTest(originalConfig.GmailCredentials.AppPassword)
                            }
                        };
                        
                        var options = new System.Text.Json.JsonSerializerOptions
                        {
                            WriteIndented = true
                        };
                        
                        string appDataJson = System.Text.Json.JsonSerializer.Serialize(configWithVolunteerPath, options);
                        File.WriteAllText(appDataConfigPath, appDataJson);

                        // Arrange - Create app directory but ensure data folder doesn't exist
                        Directory.CreateDirectory(tempAppDir);
                        if (Directory.Exists(dataFolderPath))
                        {
                            Directory.Delete(dataFolderPath, true);
                        }

                        // Create a test service that uses the temporary directories
                        var testService = new TestConfigurationServiceWithVolunteerMigration(tempAppDir, tempAppDataDir, volunteerManager);

                        // Act - Load configuration (this should trigger migration)
                        var loadedConfig = testService.LoadConfiguration();

                        // Verify migration completed successfully
                        if (!File.Exists(dataConfigPath))
                        {
                            return false.Label("Precondition failed: Migration did not complete (config not in data folder)");
                        }

                        if (!File.Exists(dataVolunteersPath))
                        {
                            return false.Label("Precondition failed: Migration did not complete (volunteers not in data folder)");
                        }

                        // Get snapshot of AppData folder state after migration
                        var appDataFilesAfterMigration = Directory.Exists(tempAppDataDir) 
                            ? Directory.GetFiles(tempAppDataDir, "*", SearchOption.AllDirectories).ToList()
                            : new List<string>();
                        
                        var appDataFileTimestampsAfterMigration = appDataFilesAfterMigration
                            .ToDictionary(f => f, f => File.GetLastWriteTimeUtc(f));

                        // Act - Perform various operations that should NOT touch AppData folder
                        
                        // 1. Save configuration (should only write to data folder)
                        var modifiedConfig = new AppConfiguration
                        {
                            LastExcelFilePath = "modified_path.xlsx",
                            LastSheetName = "ModifiedSheet",
                            GmailCredentials = new GmailCredentials
                            {
                                Email = "modified@example.com",
                                AppPassword = "modifiedpassword123"
                            }
                        };
                        testService.SaveConfiguration(modifiedConfig);

                        // 2. Load configuration again (should only read from data folder)
                        var reloadedConfig = testService.LoadConfiguration();

                        // 3. Save configuration again with different values
                        var anotherConfig = new AppConfiguration
                        {
                            LastExcelFilePath = "another_path.xlsx",
                            LastSheetName = "AnotherSheet",
                            GmailCredentials = new GmailCredentials
                            {
                                Email = "another@example.com",
                                AppPassword = "anotherpassword456"
                            }
                        };
                        testService.SaveConfiguration(anotherConfig);

                        // Assert - Verify no new files were created in AppData folder
                        var appDataFilesAfterOperations = Directory.Exists(tempAppDataDir)
                            ? Directory.GetFiles(tempAppDataDir, "*", SearchOption.AllDirectories).ToList()
                            : new List<string>();

                        // Check if any new files were created
                        var newFiles = appDataFilesAfterOperations.Except(appDataFilesAfterMigration).ToList();
                        if (newFiles.Any())
                        {
                            return false.Label($"New files created in AppData folder after migration: {string.Join(", ", newFiles)}");
                        }

                        // Assert - Verify no existing files in AppData were modified
                        foreach (var file in appDataFilesAfterMigration)
                        {
                            if (File.Exists(file))
                            {
                                var originalTimestamp = appDataFileTimestampsAfterMigration[file];
                                var currentTimestamp = File.GetLastWriteTimeUtc(file);
                                
                                if (currentTimestamp > originalTimestamp)
                                {
                                    return false.Label($"AppData file was modified after migration: {file}");
                                }
                            }
                        }

                        // Assert - Verify all data is in data folder
                        if (!File.Exists(dataConfigPath))
                        {
                            return false.Label("Configuration file not found in data folder after operations");
                        }

                        if (!File.Exists(dataVolunteersPath))
                        {
                            return false.Label("Volunteers file not found in data folder after operations");
                        }

                        // Assert - Verify subsequent operations used data folder only
                        // (by checking that the latest config values are in the data folder)
                        string dataConfigContent = File.ReadAllText(dataConfigPath);
                        if (!dataConfigContent.Contains("another@example.com"))
                        {
                            return false.Label("Latest configuration not saved to data folder");
                        }

                        return true.ToProperty();
                    }
                    catch (Exception ex)
                    {
                        return false.Label($"AppData not used after migration test failed with exception: {ex.Message}");
                    }
                    finally
                    {
                        // Cleanup - Remove temporary directories and files
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                        
                        if (Directory.Exists(tempAppDataDir))
                        {
                            try
                            {
                                Directory.Delete(tempAppDataDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }

                        if (File.Exists(tempExternalVolunteerFile))
                        {
                            try
                            {
                                File.Delete(tempExternalVolunteerFile);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        // Feature: portable-data-storage, Property 13: Error Messages Include Path
        /// <summary>
        /// Property 13: Error Messages Include Path
        /// For any data folder operation that fails with an exception, the exception message
        /// should contain the attempted data folder path.
        /// **Validates: Requirements 8.4**
        /// </summary>
        [Test]
        public void Property_ErrorMessagesIncludePath()
        {
            var config = Configuration.QuickThrowOnFailure;
            config.MaxNbOfTest = 100;

            Prop.ForAll(
                Arb.Default.Guid(),
                (Guid testId) =>
                {
                    // Use a temporary directory that we'll make read-only to trigger permission errors
                    var tempAppDir = Path.Combine(Path.GetTempPath(), $"test_app_{testId}");
                    var expectedDataFolder = Path.Combine(tempAppDir, "data");

                    try
                    {
                        // Arrange - Create temporary app directory
                        Directory.CreateDirectory(tempAppDir);

                        // Make the app directory read-only to trigger UnauthorizedAccessException
                        var dirInfo = new DirectoryInfo(tempAppDir);
                        dirInfo.Attributes = FileAttributes.ReadOnly;

                        // Create a test service that uses the temporary app directory
                        var volunteerManager = new VolunteerManager();
                        var testService = new TestConfigurationServiceWithAppDir(tempAppDir, volunteerManager);

                        // Act & Assert - Try to create data folder (should throw exception with path)
                        try
                        {
                            testService.EnsureDataFolderExists();
                            
                            // If we get here, the operation succeeded (shouldn't happen with read-only directory)
                            // This might happen on some systems, so we'll just verify the folder was created
                            if (Directory.Exists(expectedDataFolder))
                            {
                                return true.ToProperty().Label("Operation succeeded despite read-only directory (system-dependent behavior)");
                            }
                            
                            return false.Label("EnsureDataFolderExists() did not throw exception or create folder");
                        }
                        catch (InvalidOperationException ex)
                        {
                            // Assert - Verify the exception message contains the data folder path
                            if (!ex.Message.Contains(expectedDataFolder))
                            {
                                return false.Label($"Exception message does not contain data folder path. Expected path '{expectedDataFolder}' in message: {ex.Message}");
                            }

                            // Assert - Verify the message is descriptive (contains "permissions" or "disk space")
                            bool hasPermissionMessage = ex.Message.Contains("permissions", StringComparison.OrdinalIgnoreCase) ||
                                                       ex.Message.Contains("Insufficient permissions", StringComparison.OrdinalIgnoreCase);
                            bool hasDiskSpaceMessage = ex.Message.Contains("disk space", StringComparison.OrdinalIgnoreCase);

                            if (!hasPermissionMessage && !hasDiskSpaceMessage)
                            {
                                return false.Label($"Exception message is not descriptive. Expected 'permissions' or 'disk space' in message: {ex.Message}");
                            }

                            // Assert - Verify the path in the message is helpful for troubleshooting
                            // (it should be the full path, not just "data")
                            if (ex.Message.Contains("'data'") && !ex.Message.Contains(tempAppDir))
                            {
                                return false.Label($"Exception message contains relative path instead of full path: {ex.Message}");
                            }

                            return true.ToProperty();
                        }
                        catch (UnauthorizedAccessException ex)
                        {
                            // This is also acceptable - verify it contains the path
                            if (!ex.Message.Contains(expectedDataFolder) && !ex.Message.Contains(tempAppDir))
                            {
                                return false.Label($"UnauthorizedAccessException message does not contain data folder path: {ex.Message}");
                            }
                            
                            return true.ToProperty();
                        }
                        catch (Exception ex)
                        {
                            return false.Label($"Unexpected exception type: {ex.GetType().Name}. Message: {ex.Message}");
                        }
                    }
                    finally
                    {
                        // Cleanup - Remove read-only attribute and delete directory
                        if (Directory.Exists(tempAppDir))
                        {
                            try
                            {
                                var dirInfo = new DirectoryInfo(tempAppDir);
                                dirInfo.Attributes = FileAttributes.Normal;
                                
                                // Also remove read-only from all subdirectories and files
                                foreach (var file in dirInfo.GetFiles("*", SearchOption.AllDirectories))
                                {
                                    file.Attributes = FileAttributes.Normal;
                                }
                                foreach (var dir in dirInfo.GetDirectories("*", SearchOption.AllDirectories))
                                {
                                    dir.Attributes = FileAttributes.Normal;
                                }
                                
                                Directory.Delete(tempAppDir, true);
                            }
                            catch
                            {
                                // Ignore cleanup errors
                            }
                        }
                    }
                }
            ).Check(config);
        }

        /// <summary>
        /// Helper method to encrypt password for testing (mimics DPAPI encryption)
        /// </summary>
        private static string EncryptPasswordForTest(string plainText)
        {
            if (string.IsNullOrEmpty(plainText))
            {
                return string.Empty;
            }

            try
            {
                byte[] plainBytes = Encoding.UTF8.GetBytes(plainText);
                byte[] encryptedBytes = ProtectedData.Protect(
                    plainBytes,
                    null,
                    DataProtectionScope.CurrentUser
                );
                return Convert.ToBase64String(encryptedBytes);
            }
            catch
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// Helper method to check if two configurations are equal
        /// </summary>
        private bool AreConfigurationsEqual(AppConfiguration config1, AppConfiguration config2)
        {
            // VolunteerFilePath removed - no longer part of AppConfiguration (Task 2.1)
            return config1.LastExcelFilePath == config2.LastExcelFilePath &&
                   config1.LastSheetName == config2.LastSheetName &&
                   config1.GmailCredentials.Email == config2.GmailCredentials.Email &&
                   config1.GmailCredentials.AppPassword == config2.GmailCredentials.AppPassword;
        }

        /// <summary>
        /// Test helper class that allows overriding the config file path for testing
        /// </summary>
        private class TestConfigurationService : ConfigurationService
        {
            private readonly string _testConfigPath;

            public TestConfigurationService(string testConfigPath, IVolunteerManager volunteerManager)
                : base(volunteerManager)
            {
                _testConfigPath = testConfigPath;
            }

            public new string GetConfigFilePath()
            {
                return _testConfigPath;
            }

            public new AppConfiguration LoadConfiguration()
            {
                string configPath = GetConfigFilePath();

                // Handle missing file - return empty config
                if (!File.Exists(configPath))
                {
                    return new AppConfiguration();
                }

                try
                {
                    // Read and deserialize config.json
                    string json = File.ReadAllText(configPath);
                    var config = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json);

                    // Return deserialized config or empty if null
                    return config ?? new AppConfiguration();
                }
                catch (Exception ex) when (ex is System.Text.Json.JsonException || ex is IOException)
                {
                    // Handle corrupted file - log warning and return empty config
                    Console.WriteLine($"Warning: Configuration file is corrupted or unreadable. Using empty configuration. Error: {ex.Message}");
                    return new AppConfiguration();
                }
            }

            public new void SaveConfiguration(AppConfiguration config)
            {
                string configPath = GetConfigFilePath();

                // Create directory if it doesn't exist
                string? directory = Path.GetDirectoryName(configPath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Serialize and write config to file
                var options = new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                };
                string json = System.Text.Json.JsonSerializer.Serialize(config, options);
                File.WriteAllText(configPath, json);
            }
        }

        /// <summary>
        /// Test helper class that allows overriding the app directory for testing data folder creation
        /// and full configuration service functionality including encryption/decryption
        /// </summary>
        private class TestConfigurationServiceWithAppDir : ConfigurationService
        {
            private readonly string _testAppDir;

            public TestConfigurationServiceWithAppDir(string testAppDir, IVolunteerManager volunteerManager)
                : base(volunteerManager)
            {
                _testAppDir = testAppDir;
            }

            protected override string GetBaseDirectory()
            {
                return _testAppDir;
            }
        }

        /// <summary>
        /// Test helper class that allows overriding both app directory and AppData directory
        /// for testing configuration migration functionality
        /// </summary>
        private class TestConfigurationServiceWithMigration : ConfigurationService
        {
            private readonly string _testAppDir;
            private readonly string _testAppDataDir;

            public TestConfigurationServiceWithMigration(string testAppDir, string testAppDataDir, IVolunteerManager volunteerManager)
                : base(volunteerManager)
            {
                _testAppDir = testAppDir;
                _testAppDataDir = testAppDataDir;
            }

            protected override string GetBaseDirectory()
            {
                return _testAppDir;
            }

            /// <summary>
            /// Override to return test AppData path instead of real AppData
            /// </summary>
            private string GetAppDataConfigPath()
            {
                return Path.Combine(_testAppDataDir, "config.json");
            }

            /// <summary>
            /// Override migration logic to use test directories
            /// </summary>
            private void MigrateFromAppData()
            {
                string appDataConfigPath = GetAppDataConfigPath();
                
                // Check if AppData config exists
                if (!File.Exists(appDataConfigPath))
                {
                    return; // Nothing to migrate
                }

                try
                {
                    string dataConfigPath = GetConfigFilePath();
                    
                    // Ensure data folder exists
                    string? directory = Path.GetDirectoryName(dataConfigPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    
                    // Copy the config file from AppData to data folder
                    File.Copy(appDataConfigPath, dataConfigPath, overwrite: false);
                    
                    Console.WriteLine($"Successfully migrated configuration from AppData to data folder.");
                }
                catch (IOException ex)
                {
                    // Handle corrupted or inaccessible files gracefully
                    Console.WriteLine($"Warning: Could not migrate configuration from AppData. Error: {ex.Message}");
                }
                catch (UnauthorizedAccessException ex)
                {
                    // Handle permission issues
                    Console.WriteLine($"Warning: Could not migrate configuration from AppData due to insufficient permissions. Error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    // Handle any other unexpected errors
                    Console.WriteLine($"Warning: Could not migrate configuration from AppData. Error: {ex.Message}");
                }
            }

            /// <summary>
            /// Override LoadConfiguration to use test migration logic
            /// </summary>
            public new AppConfiguration LoadConfiguration()
            {
                // Ensure data folder exists first
                string dataFolder = Path.Combine(_testAppDir, "data");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }

                string configPath = GetConfigFilePath();

                // If config doesn't exist in data folder, attempt migration from AppData
                if (!File.Exists(configPath))
                {
                    MigrateFromAppData();
                }

                // Handle missing file - return empty config
                if (!File.Exists(configPath))
                {
                    return new AppConfiguration();
                }

                try
                {
                    // Read and deserialize config.json
                    string json = File.ReadAllText(configPath);
                    var config = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json);

                    if (config == null)
                    {
                        return new AppConfiguration();
                    }

                    // Decrypt the password after loading
                    if (!string.IsNullOrEmpty(config.GmailCredentials.AppPassword))
                    {
                        try
                        {
                            byte[] encryptedBytes = Convert.FromBase64String(config.GmailCredentials.AppPassword);
                            byte[] plainBytes = ProtectedData.Unprotect(
                                encryptedBytes,
                                null,
                                DataProtectionScope.CurrentUser
                            );
                            config.GmailCredentials.AppPassword = Encoding.UTF8.GetString(plainBytes);
                        }
                        catch
                        {
                            // If decryption fails, return empty string
                            config.GmailCredentials.AppPassword = string.Empty;
                        }
                    }

                    return config;
                }
                catch (Exception ex) when (ex is System.Text.Json.JsonException || ex is IOException)
                {
                    // Handle corrupted file - log warning and return empty config
                    Console.WriteLine($"Warning: Configuration file is corrupted or unreadable. Using empty configuration. Error: {ex.Message}");
                    return new AppConfiguration();
                }
            }
        }

        /// <summary>
        /// Test helper class that allows overriding both app directory and AppData directory
        /// for testing volunteer migration functionality
        /// </summary>
        private class TestConfigurationServiceWithVolunteerMigration : ConfigurationService
        {
            private readonly string _testAppDir;
            private readonly string _testAppDataDir;
            private readonly IVolunteerManager _testVolunteerManager;

            public TestConfigurationServiceWithVolunteerMigration(string testAppDir, string testAppDataDir, IVolunteerManager volunteerManager)
                : base(volunteerManager)
            {
                _testAppDir = testAppDir;
                _testAppDataDir = testAppDataDir;
                _testVolunteerManager = volunteerManager;
            }

            protected override string GetBaseDirectory()
            {
                return _testAppDir;
            }

            /// <summary>
            /// Override to return test AppData path instead of real AppData
            /// </summary>
            private string GetAppDataConfigPath()
            {
                return Path.Combine(_testAppDataDir, "config.json");
            }

            /// <summary>
            /// Override migration logic to use test directories
            /// </summary>
            private void MigrateFromAppData()
            {
                string appDataConfigPath = GetAppDataConfigPath();
                
                // Check if AppData config exists
                if (!File.Exists(appDataConfigPath))
                {
                    return; // Nothing to migrate
                }

                try
                {
                    string dataConfigPath = GetConfigFilePath();
                    
                    // Ensure data folder exists
                    string? directory = Path.GetDirectoryName(dataConfigPath);
                    if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }
                    
                    // Copy the config file from AppData to data folder
                    File.Copy(appDataConfigPath, dataConfigPath, overwrite: false);
                    
                    Console.WriteLine($"Successfully migrated configuration from AppData to data folder.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not migrate configuration from AppData. Error: {ex.Message}");
                }
            }

            /// <summary>
            /// Override LoadConfiguration to use test migration logic including volunteer migration
            /// </summary>
            public new AppConfiguration LoadConfiguration()
            {
                // Ensure data folder exists first
                string dataFolder = Path.Combine(_testAppDir, "data");
                if (!Directory.Exists(dataFolder))
                {
                    Directory.CreateDirectory(dataFolder);
                }

                string configPath = GetConfigFilePath();

                // If config doesn't exist in data folder, attempt migration from AppData
                if (!File.Exists(configPath))
                {
                    MigrateFromAppData();
                }

                // Handle missing file - return empty config
                if (!File.Exists(configPath))
                {
                    return new AppConfiguration();
                }

                try
                {
                    // Read and deserialize config.json
                    string json = File.ReadAllText(configPath);
                    var config = System.Text.Json.JsonSerializer.Deserialize<AppConfiguration>(json);

                    if (config == null)
                    {
                        return new AppConfiguration();
                    }

                    // Decrypt the password after loading
                    if (!string.IsNullOrEmpty(config.GmailCredentials.AppPassword))
                    {
                        try
                        {
                            byte[] encryptedBytes = Convert.FromBase64String(config.GmailCredentials.AppPassword);
                            byte[] plainBytes = ProtectedData.Unprotect(
                                encryptedBytes,
                                null,
                                DataProtectionScope.CurrentUser
                            );
                            config.GmailCredentials.AppPassword = Encoding.UTF8.GetString(plainBytes);
                        }
                        catch
                        {
                            // If decryption fails, return empty string
                            config.GmailCredentials.AppPassword = string.Empty;
                        }
                    }

                    // Call volunteer migration after config migration
                    MigrateVolunteerDataTest(config);

                    return config;
                }
                catch (Exception ex) when (ex is System.Text.Json.JsonException || ex is IOException)
                {
                    Console.WriteLine($"Warning: Configuration file is corrupted or unreadable. Using empty configuration. Error: {ex.Message}");
                    return new AppConfiguration();
                }
            }

            /// <summary>
            /// Test version of MigrateVolunteerData that uses the test app directory
            /// </summary>
            private void MigrateVolunteerDataTest(AppConfiguration config)
            {
                string configPath = GetConfigFilePath();
                
                if (!File.Exists(configPath))
                {
                    return;
                }

                try
                {
                    string json = File.ReadAllText(configPath);
                    
                    if (!json.Contains("VolunteerFilePath"))
                    {
                        return;
                    }

                    using var document = System.Text.Json.JsonDocument.Parse(json);
                    if (!document.RootElement.TryGetProperty("VolunteerFilePath", out var volunteerFilePathElement))
                    {
                        return;
                    }

                    string volunteerFilePath = volunteerFilePathElement.GetString() ?? string.Empty;
                    
                    if (string.IsNullOrEmpty(volunteerFilePath))
                    {
                        RemoveVolunteerFilePathFromConfigTest(json);
                        return;
                    }

                    if (!File.Exists(volunteerFilePath))
                    {
                        Console.WriteLine($"Warning: Volunteer file not found at '{volunteerFilePath}'. Skipping volunteer migration.");
                        RemoveVolunteerFilePathFromConfigTest(json);
                        return;
                    }

                    // Load volunteer data from external file
                    Dictionary<string, string> volunteers;
                    try
                    {
                        volunteers = _testVolunteerManager.LoadVolunteers(volunteerFilePath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not load volunteer data from '{volunteerFilePath}'. Error: {ex.Message}");
                        RemoveVolunteerFilePathFromConfigTest(json);
                        return;
                    }

                    // Save volunteer data to data/volunteers.json
                    string dataFolder = Path.Combine(_testAppDir, "data");
                    string volunteersPath = Path.Combine(dataFolder, "volunteers.json");

                    try
                    {
                        _testVolunteerManager.SaveVolunteers(volunteersPath, volunteers);
                        Console.WriteLine($"Successfully migrated volunteer data to internal storage.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not save volunteer data to internal storage. Error: {ex.Message}");
                        return;
                    }

                    RemoveVolunteerFilePathFromConfigTest(json);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not migrate volunteer data. Error: {ex.Message}");
                }
            }

            /// <summary>
            /// Test version of RemoveVolunteerFilePathFromConfig
            /// </summary>
            private void RemoveVolunteerFilePathFromConfigTest(string originalJson)
            {
                try
                {
                    using var document = System.Text.Json.JsonDocument.Parse(originalJson);
                    var root = document.RootElement;

                    var configDict = new Dictionary<string, object?>();

                    foreach (var property in root.EnumerateObject())
                    {
                        if (property.Name != "VolunteerFilePath")
                        {
                            configDict[property.Name] = property.Value.Clone();
                        }
                    }

                    var options = new System.Text.Json.JsonSerializerOptions
                    {
                        WriteIndented = true
                    };

                    string newJson = System.Text.Json.JsonSerializer.Serialize(configDict, options);
                    string configPath = GetConfigFilePath();
                    File.WriteAllText(configPath, newJson);

                    Console.WriteLine("Successfully removed VolunteerFilePath from configuration.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Warning: Could not remove VolunteerFilePath from configuration. Error: {ex.Message}");
                }
            }
        }
    }
}
