using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Services;

/// <summary>
/// Orchestrates the volunteer notification workflow.
/// Coordinates between volunteer management, email sending, configuration persistence,
/// Excel data reading, and UI updates.
/// </summary>
public class VolunteerNotificationController : IVolunteerNotificationController
{
    private readonly IVolunteerManager _volunteerManager;
    private readonly IEmailService _emailService;
    private readonly IConfigurationService _configurationService;
    private readonly IExcelManager _excelManager;
    private readonly IVolunteerUI _ui;
    private readonly string _dataFolderPath;

    private Dictionary<string, string> _volunteers;
    private GmailCredentials _gmailCredentials;
    private string _volunteerFilePath;
    private string _selectedExcelFilePath;
    private string _selectedSheetName;

    /// <summary>
    /// Initializes a new instance of the VolunteerNotificationController.
    /// Loads configuration from persistent storage on construction.
    /// </summary>
    /// <param name="volunteerManager">Service for managing volunteer contacts</param>
    /// <param name="emailService">Service for sending emails via Gmail SMTP</param>
    /// <param name="configurationService">Service for persistent configuration storage</param>
    /// <param name="excelManager">Service for reading Excel files</param>
    /// <param name="ui">UI interface for displaying data and messages</param>
    /// <param name="dataFolderPath">Optional custom data folder path for testing. If null, uses AppDomain.CurrentDomain.BaseDirectory/data</param>
    public VolunteerNotificationController(
        IVolunteerManager volunteerManager,
        IEmailService emailService,
        IConfigurationService configurationService,
        IExcelManager excelManager,
        IVolunteerUI ui,
        string dataFolderPath = null)
    {
        _volunteerManager = volunteerManager;
        _emailService = emailService;
        _configurationService = configurationService;
        _excelManager = excelManager;
        _ui = ui;
        _dataFolderPath = dataFolderPath ?? Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "data");

        // Initialize state fields
        _volunteers = new Dictionary<string, string>();
        _gmailCredentials = new GmailCredentials();
        _volunteerFilePath = string.Empty;
        _selectedExcelFilePath = string.Empty;
        _selectedSheetName = string.Empty;

        // Load configuration from persistent storage (Requirements 1.6, 3.3)
        LoadConfiguration();
        
        // Load volunteers from internal storage (Requirement 6.4)
        LoadVolunteersFromInternalStorage();
    }

    /// <summary>
    /// Loads configuration from persistent storage and initializes state.
    /// </summary>
    private void LoadConfiguration()
    {
        var config = _configurationService.LoadConfiguration();

        // Load Gmail credentials
        _gmailCredentials = config.GmailCredentials;

        // Load last selected Excel file and sheet
        _selectedExcelFilePath = config.LastExcelFilePath;
        _selectedSheetName = config.LastSheetName;
    }
    
    /// <summary>
    /// Refreshes the UI display with current configuration data.
    /// Should be called after the UI is fully initialized.
    /// </summary>
    public void RefreshUIDisplay()
    {
        // Display Gmail credentials in UI (Requirement 3.3)
        if (_gmailCredentials != null)
        {
            _ui.DisplayGmailCredentials(_gmailCredentials.Email ?? string.Empty, 
                                       _gmailCredentials.AppPassword ?? string.Empty);
        }
        
        // Display volunteer list in UI (Requirement 8.1)
        if (_volunteers.Count > 0)
        {
            _ui.DisplayVolunteerList(_volunteers);
        }
        
        // Update send emails button state
        _ui.EnableSendEmailsButton(CanSendEmails());
    }

    public void OnVolunteerFileSelected(string filePath)
        {
            try
            {
                // Load volunteers from the external file (Requirement 3.1)
                var importedVolunteers = _volunteerManager.LoadVolunteers(filePath);

                // Merge imported volunteers with existing data (Requirement 3.2, 3.3)
                MergeVolunteers(importedVolunteers);

                // Save merged data to internal storage (Requirement 3.4)
                SaveVolunteersToInternalStorage();

                // No external file path is stored (Requirement 3.5)

                // Display volunteers in UI (Requirement 8.1)
                _ui.DisplayVolunteerList(_volunteers);

                // Update CanSendEmails state to enable/disable send button (Requirement 5.1)
                _ui.EnableSendEmailsButton(CanSendEmails());
            }
            catch (FileNotFoundException)
            {
                // Handle file not found error with Italian message (Requirement 1.5)
                _ui.ShowErrorMessage(Properties.Resources.ErrorVolunteerFileNotFound);
            }
            catch (InvalidOperationException)
            {
                // Handle invalid JSON error with Italian message (Requirement 1.5)
                _ui.ShowErrorMessage(Properties.Resources.ErrorInvalidVolunteerFile);
            }
            catch (Exception ex)
            {
                // Handle any other errors with Italian message
                _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
            }
        }


    public void OnAddVolunteer(string surname, string email)
    {
        try
        {
            // Add volunteer to the dictionary (Requirement 8.4)
            // This will throw ArgumentException if surname is empty or email is invalid
            _volunteerManager.AddVolunteer(surname, email, _volunteers);
            
            // Save updated volunteers to file (Requirement 8.5)
            if (!string.IsNullOrEmpty(_volunteerFilePath))
            {
                _volunteerManager.SaveVolunteers(_volunteerFilePath, _volunteers);
            }
            
            // Volunteer data is now saved to internal storage (no config update needed)
            
            // Refresh UI volunteer list (Requirement 8.11)
            _ui.DisplayVolunteerList(_volunteers);
            
            // Update CanSendEmails state to enable/disable send button
            _ui.EnableSendEmailsButton(CanSendEmails());
        }
        catch (ArgumentException ex)
        {
            // Handle validation errors with Italian messages (Requirement 8.12)
            if (ex.ParamName == "surname")
            {
                _ui.ShowErrorMessage("Il cognome non può essere vuoto.");
            }
            else if (ex.ParamName == "email")
            {
                _ui.ShowErrorMessage("L'indirizzo email non è valido.");
            }
            else
            {
                _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
            }
        }
        catch (IOException ex)
        {
            // Handle file save errors
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }

    public void OnDeleteVolunteer(string surname)
        {
            try
            {
                // Remove volunteer from the dictionary (Requirement 8.10)
                _volunteerManager.RemoveVolunteer(surname, _volunteers);

                // Save updated volunteers to external file if path exists (Requirement 8.11)
                if (!string.IsNullOrEmpty(_volunteerFilePath))
                {
                    _volunteerManager.SaveVolunteers(_volunteerFilePath, _volunteers);
                }

                // Save updated volunteers to internal storage (volunteers.json)
                SaveVolunteersToInternalStorage();

                // Refresh UI volunteer list (Requirement 8.11)
                _ui.DisplayVolunteerList(_volunteers);

                // Update send emails button state (Requirement 5.1)
                _ui.EnableSendEmailsButton(CanSendEmails());
            }
            catch (IOException ex)
            {
                // Handle file save errors with Italian message
                _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
            }
            catch (Exception ex)
            {
                // Handle any other unexpected errors with Italian message
                _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
            }
        }


    public void OnDeleteAllVolunteers()
    {
        try
        {
            // Prompt user for confirmation with Italian message (Requirement 8.7)
            // Note: Using string literal temporarily until Resources.Designer.cs is regenerated
            bool confirmed = _ui.ConfirmAction("Sei sicuro di voler eliminare tutti i contatti dei volontari? Questa operazione non può essere annullata.");
            
            // If user confirms, proceed with deletion (Requirement 8.8)
            if (confirmed)
            {
                // Clear the volunteers dictionary (Requirement 8.8)
                _volunteers.Clear();
                
                // Save empty volunteers to file (Requirement 8.8)
                if (!string.IsNullOrEmpty(_volunteerFilePath))
                {
                    _volunteerManager.SaveVolunteers(_volunteerFilePath, _volunteers);
                }
                
                // Save updated volunteers to internal storage (volunteers.json)
                SaveVolunteersToInternalStorage();
                
                // Refresh UI volunteer list (Requirement 8.11)
                _ui.DisplayVolunteerList(_volunteers);
                
                // Update send emails button state (Requirement 5.1)
                _ui.EnableSendEmailsButton(CanSendEmails());
            }
        }
        catch (IOException ex)
        {
            // Handle file save errors with Italian message
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors with Italian message
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }

    public void OnGmailCredentialsUpdated(string email, string appPassword)
    {
        try
        {
            // Update internal GmailCredentials object (Requirement 3.1)
            _gmailCredentials.Email = email;
            _gmailCredentials.AppPassword = appPassword;
            
            // Save configuration to persist credentials (Requirements 3.2, 3.4)
            var config = _configurationService.LoadConfiguration();
            config.GmailCredentials = _gmailCredentials;
            _configurationService.SaveConfiguration(config);
            
            // Update CanSendEmails state to enable/disable send button (Requirement 5.1)
            _ui.EnableSendEmailsButton(CanSendEmails());
        }
        catch (Exception ex)
        {
            // Handle any errors with Italian message
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }

    public void OnClearGmailCredentials()
    {
        try
        {
            // Clear internal GmailCredentials object
            _gmailCredentials.Email = string.Empty;
            _gmailCredentials.AppPassword = string.Empty;
            
            // Save configuration to persist cleared credentials
            var config = _configurationService.LoadConfiguration();
            config.GmailCredentials = _gmailCredentials;
            _configurationService.SaveConfiguration(config);
            
            // Clear UI fields
            _ui.DisplayGmailCredentials(string.Empty, string.Empty);
            
            // Update CanSendEmails state to disable send button
            _ui.EnableSendEmailsButton(CanSendEmails());
        }
        catch (Exception ex)
        {
            // Handle any errors with Italian message
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }

    public void OnNotificationExcelFileSelected(string filePath)
    {
        try
        {
            // Open the Excel workbook to read sheet names (Requirement 2.2)
            var workbook = _excelManager.OpenWorkbook(filePath);

            // Get all sheet names from the workbook (Requirement 2.2)
            var sheetNames = _excelManager.GetSheetNames(workbook);

            // Check if the Excel file contains any sheets (Requirement 2.6)
            if (sheetNames == null || sheetNames.Count == 0)
            {
                _ui.ShowErrorMessage("Il file Excel non contiene fogli.");
                return;
            }

            // Display sheet names in the UI selection control (Requirement 2.3)
            _ui.DisplaySheetNames(sheetNames);

            // Save the Excel file path to configuration for persistence (Requirement 2.2)
            _selectedExcelFilePath = filePath;
            var config = _configurationService.LoadConfiguration();
            config.LastExcelFilePath = filePath;
            _configurationService.SaveConfiguration(config);

            // Update send emails button state (Requirement 5.1)
            _ui.EnableSendEmailsButton(CanSendEmails());
        }
        catch (FileNotFoundException)
        {
            // Handle file not found error with Italian message (Requirement 2.5)
            _ui.ShowErrorMessage("Il file Excel non è stato trovato.");
        }
        catch (InvalidOperationException)
        {
            // Handle unreadable file error with Italian message (Requirement 2.5)
            _ui.ShowErrorMessage(Properties.Resources.ErrorExcelFileRead);
        }
        catch (Exception ex)
        {
            // Handle any other errors with Italian message (Requirement 2.5)
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }


    public void OnSheetSelected(string sheetName)
    {
        try
        {
            // Store the selected sheet name (Requirement 2.4)
            _selectedSheetName = sheetName;
            
            // Save to configuration for persistence (Requirement 2.4)
            var config = _configurationService.LoadConfiguration();
            config.LastSheetName = sheetName;
            _configurationService.SaveConfiguration(config);
            
            // Update CanSendEmails state to enable/disable send button (Requirement 5.1)
            _ui.EnableSendEmailsButton(CanSendEmails());
        }
        catch (Exception ex)
        {
            // Handle any errors with Italian message
            _ui.ShowErrorMessage(string.Format(Properties.Resources.ErrorGeneral, ex.Message));
        }
    }

    public async Task OnSendEmailsClickedAsync()
    {
        try
        {
            // Validate prerequisites using CanSendEmails (Requirement 5.1)
            if (!CanSendEmails())
            {
                _ui.ShowErrorMessage("Impossibile inviare email. Verificare che le credenziali Gmail, i volontari e il foglio Excel siano configurati.");
                return;
            }

            // Open the Excel workbook
            var workbook = _excelManager.OpenWorkbook(_selectedExcelFilePath);
            
            // Get the selected sheet
            var sheet = _excelManager.GetSheetByName(workbook, _selectedSheetName);
            if (sheet == null)
            {
                _ui.ShowErrorMessage($"Il foglio '{_selectedSheetName}' non è stato trovato nel file Excel.");
                return;
            }

            // Call ExcelManager to identify volunteer assignments (Requirement 4.3)
            List<VolunteerAssignment> assignments;
            try
            {
                assignments = _excelManager.IdentifyVolunteerAssignments(sheet, _volunteers);
                
                // Debug: Show how many volunteers were found
                System.Diagnostics.Debug.WriteLine($"Found {assignments.Count} volunteers with assignments");
                foreach (var a in assignments)
                {
                    System.Diagnostics.Debug.WriteLine($"  - {a.Surname} ({a.Email}): {a.AssignedRows.Count} rows");
                }
            }
            catch (InvalidOperationException ex)
            {
                // Handle missing Volontario column error (Requirement 4.4)
                _ui.ShowErrorMessage(ex.Message);
                return;
            }

            // Track success and failure counts (Requirement 5.7)
            int successCount = 0;
            int failureCount = 0;

            // Check if any volunteers were found
            if (assignments.Count == 0)
            {
                _ui.ShowErrorMessage("Nessun volontario trovato nel foglio selezionato. Verificare che i cognomi dei volontari corrispondano a quelli nella colonna 'Volontario'.");
                return;
            }

            // For each volunteer with assignments, send email notification (Requirements 5.2, 5.3)
            foreach (var assignment in assignments)
            {
                // Skip volunteers with no assigned rows
                if (assignment.AssignedRows == null || assignment.AssignedRows.Count == 0)
                {
                    continue;
                }

                // Show progress (Requirement 5.7)
                _ui.ShowEmailProgress($"Invio email a {assignment.Surname}...");

                try
                {
                    // Call EmailService to send notification (Requirements 5.2, 5.3, 5.8)
                    bool success = await _emailService.SendVolunteerNotificationAsync(
                        assignment.Email,
                        assignment.Surname,
                        assignment.AssignedRows,
                        _gmailCredentials);

                    if (success)
                    {
                        successCount++;
                    }
                    else
                    {
                        failureCount++;
                    }
                }
                catch (Exception ex)
                {
                    // Continue on individual failures (Requirement 5.6 - resilience)
                    failureCount++;
                    // Log the error (could be enhanced with proper logging)
                    System.Diagnostics.Debug.WriteLine($"Failed to send email to {assignment.Email}: {ex.Message}");
                }
            }

            // Show summary with counts (Requirement 5.7)
            _ui.ShowEmailSummary(successCount, failureCount);
        }
        catch (FileNotFoundException)
        {
            // Handle Excel file not found error with Italian message
            _ui.ShowErrorMessage("Il file Excel non è stato trovato.");
        }
        catch (Exception ex)
        {
            // Handle any other errors with Italian message
            _ui.ShowErrorMessage($"Errore durante l'invio delle email: {ex.Message}");
        }
    }

    public Dictionary<string, string> GetVolunteers()
    {
        // Return current volunteers dictionary (Requirement 8.1)
        return _volunteers;
    }

    public (string Email, string AppPassword) GetGmailCredentials()
    {
        // Return current Gmail credentials
        return (_gmailCredentials.Email ?? string.Empty, _gmailCredentials.AppPassword ?? string.Empty);
    }

    public bool CanSendEmails()
    {
        // Check all prerequisites for sending emails (Requirements 2.1, 3.6, 5.1)
        return _gmailCredentials.IsConfigured &&
               _volunteers.Count > 0 &&
               !string.IsNullOrEmpty(_selectedSheetName);
    }

    /// <summary>
    /// Loads volunteer data from internal storage (data/volunteers.json).
    /// If the file doesn't exist, leaves the volunteers dictionary empty.
    /// </summary>
    private void LoadVolunteersFromInternalStorage()
    {
        string volunteersPath = Path.Combine(GetDataFolderPath(), "volunteers.json");
        if (File.Exists(volunteersPath))
        {
            _volunteers = _volunteerManager.LoadVolunteers(volunteersPath);
        }
    }

    /// <summary>
    /// Saves volunteer data to internal storage (data/volunteers.json).
    /// Handles file I/O errors with descriptive messages.
    /// </summary>
    private void SaveVolunteersToInternalStorage()
    {
        try
        {
            string dataFolder = GetDataFolderPath();
            string volunteersPath = Path.Combine(dataFolder, "volunteers.json");
            _volunteerManager.SaveVolunteers(volunteersPath, _volunteers);
        }
        catch (UnauthorizedAccessException ex)
        {
            throw new InvalidOperationException(
                $"Cannot save volunteer data to '{GetDataFolderPath()}'. Insufficient permissions. " +
                $"Please ensure the application has write access to this location.", ex);
        }
        catch (IOException ex)
        {
            throw new InvalidOperationException(
                $"Cannot save volunteer data to '{GetDataFolderPath()}'. " +
                $"An I/O error occurred: {ex.Message}", ex);
        }
    }

    /// <summary>
    /// Merges imported volunteer data with existing volunteers.
    /// New surnames are added, and existing surnames have their email addresses overwritten.
    /// </summary>
    /// <param name="importedVolunteers">Dictionary of imported volunteers to merge</param>
    private void MergeVolunteers(Dictionary<string, string> importedVolunteers)
    {
        // Iterate through imported volunteers (Requirements 3.2, 3.3)
        foreach (var volunteer in importedVolunteers)
        {
            // Add new surnames or overwrite existing email addresses
            _volunteers[volunteer.Key] = volunteer.Value;
        }
    }

    /// <summary>
    /// Gets the path to the data folder for storing application data.
    /// Uses the application's base directory to ensure portability.
    /// </summary>
    /// <returns>The full path to the data folder</returns>
    private string GetDataFolderPath()
    {
        return _dataFolderPath;
    }

}
