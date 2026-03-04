using AuserExcelTransformer.Models;
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Manages persistent storage of application configuration data.
/// </summary>
public class ConfigurationService : IConfigurationService
{
    /// <summary>
    /// Gets the default configuration file path.
    /// </summary>
    /// <returns>Path to config.json in user's AppData folder</returns>
    public virtual string GetConfigFilePath()
    {
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string appFolder = Path.Combine(appDataPath, "AuserExcelTransformer");
        return Path.Combine(appFolder, "config.json");
    }

    /// <summary>
    /// Loads application configuration from persistent storage.
    /// </summary>
    /// <returns>Configuration object with volunteer file path and Gmail credentials</returns>
    public AppConfiguration LoadConfiguration()
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

            if (config == null)
            {
                return new AppConfiguration();
            }

            // Decrypt the password after loading
            if (!string.IsNullOrEmpty(config.GmailCredentials.AppPassword))
            {
                config.GmailCredentials.AppPassword = DecryptPassword(config.GmailCredentials.AppPassword);
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

    /// <summary>
    /// Saves application configuration to persistent storage.
    /// </summary>
    /// <param name="config">Configuration object to save</param>
    public void SaveConfiguration(AppConfiguration config)
    {
        string configPath = GetConfigFilePath();
        string? directory = Path.GetDirectoryName(configPath);

        // Create directory if it doesn't exist
        if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
        {
            Directory.CreateDirectory(directory);
        }

        // Create a copy of the config with encrypted password
        var configToSave = new AppConfiguration
        {
            VolunteerFilePath = config.VolunteerFilePath,
            LastExcelFilePath = config.LastExcelFilePath,
            LastSheetName = config.LastSheetName,
            GmailCredentials = new GmailCredentials
            {
                Email = config.GmailCredentials.Email,
                AppPassword = EncryptPassword(config.GmailCredentials.AppPassword)
            }
        };

        // Serialize configuration to JSON
        var options = new System.Text.Json.JsonSerializerOptions
        {
            WriteIndented = true
        };
        string json = System.Text.Json.JsonSerializer.Serialize(configToSave, options);

        // Write to config.json
        File.WriteAllText(configPath, json);
    }

    /// <summary>
    /// Encrypts a password using Windows DPAPI (Data Protection API).
    /// </summary>
    /// <param name="plainText">Plain text password</param>
    /// <returns>Base64-encoded encrypted password</returns>
    private string EncryptPassword(string plainText)
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
                null, // No additional entropy
                DataProtectionScope.CurrentUser // Encrypt for current user only
            );
            return Convert.ToBase64String(encryptedBytes);
        }
        catch
        {
            // If encryption fails, return empty string
            return string.Empty;
        }
    }

    /// <summary>
    /// Decrypts a password using Windows DPAPI (Data Protection API).
    /// </summary>
    /// <param name="encryptedText">Base64-encoded encrypted password</param>
    /// <returns>Decrypted plain text password</returns>
    private string DecryptPassword(string encryptedText)
    {
        if (string.IsNullOrEmpty(encryptedText))
        {
            return string.Empty;
        }

        try
        {
            byte[] encryptedBytes = Convert.FromBase64String(encryptedText);
            byte[] plainBytes = ProtectedData.Unprotect(
                encryptedBytes,
                null, // No additional entropy
                DataProtectionScope.CurrentUser // Decrypt for current user only
            );
            return Encoding.UTF8.GetString(plainBytes);
        }
        catch
        {
            // If decryption fails, return empty string
            return string.Empty;
        }
    }

}
