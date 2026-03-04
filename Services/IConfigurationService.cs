using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Manages persistent storage of application configuration data.
/// </summary>
public interface IConfigurationService
{
    /// <summary>
    /// Loads application configuration from persistent storage.
    /// </summary>
    /// <returns>Configuration object with volunteer file path and Gmail credentials</returns>
    AppConfiguration LoadConfiguration();
    
    /// <summary>
    /// Saves application configuration to persistent storage.
    /// </summary>
    /// <param name="config">Configuration object to save</param>
    void SaveConfiguration(AppConfiguration config);
    
    /// <summary>
    /// Gets the default configuration file path.
    /// </summary>
    /// <returns>Path to config.json in user's AppData folder</returns>
    string GetConfigFilePath();
}
