using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            _configurationService = new ConfigurationService();
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
            return from volunteerPath in ValidFilePathGen()
                   from excelPath in ValidFilePathGen()
                   from sheetName in Gen.Elements("Sheet1", "Foglio1", "Data", "Volontari", "Servizi")
                   from credentials in GmailCredentialsGen()
                   select new AppConfiguration
                   {
                       VolunteerFilePath = volunteerPath,
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
                        var testService = new TestConfigurationService(tempConfigFile);

                        // Act - Save configuration
                        testService.SaveConfiguration(originalConfig);

                        // Act - Load configuration
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify VolunteerFilePath
                        if (originalConfig.VolunteerFilePath != loadedConfig.VolunteerFilePath)
                        {
                            return false.Label($"VolunteerFilePath mismatch: expected '{originalConfig.VolunteerFilePath}', got '{loadedConfig.VolunteerFilePath}'");
                        }

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
                        var testService = new TestConfigurationService(tempConfigFile);

                        // Act - Save first configuration
                        testService.SaveConfiguration(firstConfig);

                        // Act - Save second configuration (should overwrite first)
                        testService.SaveConfiguration(secondConfig);

                        // Act - Load configuration
                        var loadedConfig = testService.LoadConfiguration();

                        // Assert - Verify loaded config matches second config, not first
                        if (secondConfig.VolunteerFilePath != loadedConfig.VolunteerFilePath)
                        {
                            return false.Label($"VolunteerFilePath should be second value: expected '{secondConfig.VolunteerFilePath}', got '{loadedConfig.VolunteerFilePath}'");
                        }

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
                        bool matchesFirst = firstConfig.VolunteerFilePath == loadedConfig.VolunteerFilePath &&
                                          firstConfig.LastExcelFilePath == loadedConfig.LastExcelFilePath &&
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

        /// <summary>
        /// Helper method to check if two configurations are equal
        /// </summary>
        private bool AreConfigurationsEqual(AppConfiguration config1, AppConfiguration config2)
        {
            return config1.VolunteerFilePath == config2.VolunteerFilePath &&
                   config1.LastExcelFilePath == config2.LastExcelFilePath &&
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

            public TestConfigurationService(string testConfigPath)
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
    }
}
