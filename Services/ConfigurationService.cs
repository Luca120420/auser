using AuserExcelTransformer.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Manages persistent storage of application configuration data.
/// </summary>
public class ConfigurationService : IConfigurationService
{
    private readonly IVolunteerManager _volunteerManager;

    /// <summary>
    /// Initializes a new instance of the ConfigurationService class.
    /// </summary>
    /// <param name="volunteerManager">Volunteer manager for handling volunteer data migration</param>
    public ConfigurationService(IVolunteerManager volunteerManager)
    {
        _volunteerManager = volunteerManager ?? throw new ArgumentNullException(nameof(volunteerManager));
    }

    /// <summary>
    /// Gets the application base directory. Virtual to allow overriding in tests.
    /// </summary>
    /// <returns>The application base directory path</returns>
    protected virtual string GetBaseDirectory()
    {
        return AppDomain.CurrentDomain.BaseDirectory;
    }

    /// <summary>
    /// Gets the default configuration file path.
    /// </summary>
    /// <returns>Path to config.json in the application's data folder</returns>
    public virtual string GetConfigFilePath()
    {
        string appFolder = GetBaseDirectory();
        string dataFolder = Path.Combine(appFolder, "data");
        return Path.Combine(dataFolder, "config.json");
    }

    /// <summary>
    /// Ensures the data folder exists, creating it if necessary.
    /// </summary>
    /// <exception cref="InvalidOperationException">Thrown when the data folder cannot be created due to insufficient permissions or disk space issues.</exception>
    public void EnsureDataFolderExists()
    {
        string appFolder = GetBaseDirectory();
        string dataFolder = Path.Combine(appFolder, "data");

        // If folder already exists, nothing to do
        if (Directory.Exists(dataFolder))
        {
            return;
        }

        try
        {
            Directory.CreateDirectory(dataFolder);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new InvalidOperationException(
                $"Cannot create data folder at '{dataFolder}'. Insufficient permissions. " +
                $"Please ensure the application has write access to this location.", ex);
        }
        catch (IOException ex) when (ex.HResult == -2147024784) // Disk full (0x80070070)
        {
            throw new InvalidOperationException(
                $"Cannot create data folder at '{dataFolder}'. Insufficient disk space.", ex);
        }
    }

    /// <summary>
    /// Gets the legacy AppData configuration file path for migration purposes only.
    /// </summary>
    /// <returns>Path to config.json in the AppData folder</returns>
    private string GetAppDataConfigPath()
    {
        string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
        string appFolder = Path.Combine(appDataPath, "AuserExcelTransformer");
        return Path.Combine(appFolder, "config.json");
    }

    /// <summary>
    /// Migrates configuration from the legacy AppData folder to the data folder.
    /// This method is called only when the data folder config doesn't exist.
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
    /// Migrates volunteer data from external file to internal storage.
    /// Checks if the migrated configuration contains a VolunteerFilePath property,
    /// loads the volunteer data, saves it to data/volunteers.json, and removes the property.
    /// </summary>
    /// <param name="config">The migrated configuration that may contain VolunteerFilePath</param>
    private void MigrateVolunteerData(AppConfiguration config)
    {
        string configPath = GetConfigFilePath();
        
        // Read the raw JSON to check for VolunteerFilePath property
        if (!File.Exists(configPath))
        {
            return; // No config file to check
        }

        try
        {
            string json = File.ReadAllText(configPath);
            
            // Check if VolunteerFilePath exists in the JSON
            if (!json.Contains("VolunteerFilePath"))
            {
                return; // No volunteer file path to migrate
            }

            // Parse JSON to extract VolunteerFilePath
            using var document = System.Text.Json.JsonDocument.Parse(json);
            if (!document.RootElement.TryGetProperty("VolunteerFilePath", out var volunteerFilePathElement))
            {
                return; // Property doesn't exist
            }

            string volunteerFilePath = volunteerFilePathElement.GetString() ?? string.Empty;
            
            if (string.IsNullOrEmpty(volunteerFilePath))
            {
                // Remove the property even if it's empty
                RemoveVolunteerFilePathFromConfig(json);
                return;
            }

            // Check if external volunteer file exists
            if (!File.Exists(volunteerFilePath))
            {
                Console.WriteLine($"Warning: Volunteer file not found at '{volunteerFilePath}'. Skipping volunteer migration.");
                // Remove the property since the file doesn't exist
                RemoveVolunteerFilePathFromConfig(json);
                return;
            }

            // Load volunteer data from external file
            Dictionary<string, string> volunteers;
            try
            {
                volunteers = _volunteerManager.LoadVolunteers(volunteerFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not load volunteer data from '{volunteerFilePath}'. Error: {ex.Message}");
                // Remove the property since we can't load the data
                RemoveVolunteerFilePathFromConfig(json);
                return;
            }

            // Save volunteer data to data/volunteers.json
            string appFolder = GetBaseDirectory();
            string dataFolder = Path.Combine(appFolder, "data");
            string volunteersPath = Path.Combine(dataFolder, "volunteers.json");

            try
            {
                _volunteerManager.SaveVolunteers(volunteersPath, volunteers);
                Console.WriteLine($"Successfully migrated volunteer data to internal storage.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not save volunteer data to internal storage. Error: {ex.Message}");
                return; // Don't remove the property if we couldn't save
            }

            // Remove VolunteerFilePath property from configuration
            RemoveVolunteerFilePathFromConfig(json);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not migrate volunteer data. Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Removes the VolunteerFilePath property from the configuration file.
    /// </summary>
    /// <param name="originalJson">The original JSON content</param>
    private void RemoveVolunteerFilePathFromConfig(string originalJson)
    {
        try
        {
            string configPath = GetConfigFilePath();
            
            // Read the current file content
            string currentJson = File.ReadAllText(configPath);
            
            // Parse the JSON to extract properties
            using var document = System.Text.Json.JsonDocument.Parse(currentJson);
            var root = document.RootElement;
            
            // Build a new JSON object without VolunteerFilePath
            using var stream = new MemoryStream();
            using (var writer = new System.Text.Json.Utf8JsonWriter(stream, new System.Text.Json.JsonWriterOptions { Indented = true }))
            {
                writer.WriteStartObject();
                
                foreach (var property in root.EnumerateObject())
                {
                    if (property.Name != "VolunteerFilePath")
                    {
                        property.WriteTo(writer);
                    }
                }
                
                writer.WriteEndObject();
            }
            
            // Write the cleaned JSON back to the file
            string cleanedJson = Encoding.UTF8.GetString(stream.ToArray());
            File.WriteAllText(configPath, cleanedJson);
            
            Console.WriteLine("Removed VolunteerFilePath property from configuration.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Could not remove VolunteerFilePath from configuration. Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Loads application configuration from persistent storage.
    /// Ensures data folder exists and migrates from AppData if necessary.
    /// </summary>
    /// <returns>Configuration object with volunteer file path and Gmail credentials</returns>
    public AppConfiguration LoadConfiguration()
    {
        // Ensure data folder exists first
        EnsureDataFolderExists();

        string configPath = GetConfigFilePath();

        // If config doesn't exist in data folder, attempt migration from AppData
        if (!File.Exists(configPath))
        {
            MigrateFromAppData();
        }

        // Load configuration
        var config = LoadConfigurationInternal();
        
        // Migrate volunteer data after config migration
        MigrateVolunteerData(config);

        // Reload configuration after volunteer migration (in case it was cleaned)
        return LoadConfigurationInternal();
    }

    /// <summary>
    /// Internal method to load configuration without triggering migration.
    /// </summary>
    /// <returns>Configuration object</returns>
    private AppConfiguration LoadConfigurationInternal()
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
