using System;
using System.IO;
using NUnit.Framework;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for ConfigurationService class.
    /// Tests specific examples, edge cases, and security requirements.
    /// </summary>
    [TestFixture]
    public class ConfigurationServiceTests
    {
        private ConfigurationService _configurationService = null!;

        [SetUp]
        public void Setup()
        {
            _configurationService = new ConfigurationService();
        }

        /// <summary>
        /// Property 6: Password Secure Storage
        /// For any Gmail application password, storing it in persistent storage
        /// should not result in the password being stored in plain text format.
        /// **Validates: Requirements 3.5**
        /// </summary>
        [Test]
        public void SaveConfiguration_ShouldNotStorePlainTextPassword()
        {
            // Arrange
            var tempDir = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDir);
            var tempConfigFile = Path.Combine(tempDir, "config.json");

            // Use reflection to temporarily override the config file path
            var testPassword = "MySecretPassword123!";
            var config = new AppConfiguration
            {
                VolunteerFilePath = "volunteers.json",
                GmailCredentials = new GmailCredentials
                {
                    Email = "test@example.com",
                    AppPassword = testPassword
                }
            };

            try
            {
                // Act - Save configuration using a custom service that uses the temp path
                var service = new TestableConfigurationService(tempConfigFile);
                service.SaveConfiguration(config);

                // Assert - Read the raw file content
                string fileContent = File.ReadAllText(tempConfigFile);

                // Verify the password is NOT stored in plain text
                Assert.That(fileContent, Does.Not.Contain(testPassword),
                    "Password should not be stored in plain text in the configuration file. " +
                    "The password must be encrypted or hashed before storage to meet security requirement 3.5.");
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

        /// <summary>
        /// Test helper class that allows overriding the config file path for testing
        /// </summary>
        private class TestableConfigurationService : ConfigurationService
        {
            private readonly string _testConfigPath;

            public TestableConfigurationService(string testConfigPath)
            {
                _testConfigPath = testConfigPath;
            }

            // Override GetConfigFilePath to return test path
            public override string GetConfigFilePath()
            {
                return _testConfigPath;
            }
        }
    }
}
