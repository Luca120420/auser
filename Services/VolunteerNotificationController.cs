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
    public VolunteerNotificationController(
        IVolunteerManager volunteerManager,
        IEmailService emailService,
        IConfigurationService configurationService,
        IExcelManager excelManager,
        IVolunteerUI ui)
    {
        _volunteerManager = volunteerManager;
        _emailService = emailService;
        _configurationService = configurationService;
        _excelManager = excelManager;
        _ui = ui;

        // Initialize state fields
        _volunteers = new Dictionary<string, string>();
        _gmailCredentials = new GmailCredentials();
        _volunteerFilePath = string.Empty;
        _selectedExcelFilePath = string.Empty;
        _selectedSheetName = string.Empty;

        // Load configuration from persistent storage (Requirements 1.6, 3.3)
        LoadConfiguration();
    }

    /// <summary>
    /// Loads configuration from persistent storage and initializes state.
    /// </summary>
    private void LoadConfiguration()
    {
        var config = _configurationService.LoadConfiguration();

        // Load Gmail credentials
        _gmailCredentials = config.GmailCredentials;

        // Load volunteer file if path exists
        if (!string.IsNullOrEmpty(config.VolunteerFilePath))
        {
            _volunteerFilePath = config.VolunteerFilePath;
            try
            {
                _volunteers = _volunteerManager.LoadVolunteers(_volunteerFilePath);
            }
            catch
            {
                // If loading fails, start with empty volunteers
                // Error will be shown when user tries to use the feature
                _volunteers = new Dictionary<string, string>();
            }
        }

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
            // Load volunteers from the selected file (Requirement 1.4)
            _volunteers = _volunteerManager.LoadVolunteers(filePath);
            
            // Update internal state with the file path (Requirement 1.3)
            _volunteerFilePath = filePath;
            
            // Save file path to configuration for persistence (Requirement 1.3, 7.1)
            var config = _configurationService.LoadConfiguration();
            config.VolunteerFilePath = filePath;
            _configurationService.SaveConfiguration(config);
            
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
            
            // Update configuration (persist changes)
            var config = _configurationService.LoadConfiguration();
            config.VolunteerFilePath = _volunteerFilePath;
            _configurationService.SaveConfiguration(config);
            
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
            
            // Save updated volunteers to file (Requirement 8.11)
            if (!string.IsNullOrEmpty(_volunteerFilePath))
            {
                _volunteerManager.SaveVolunteers(_volunteerFilePath, _volunteers);
            }
            
            // Update configuration (persist changes)
            var config = _configurationService.LoadConfiguration();
            config.VolunteerFilePath = _volunteerFilePath;
            _configurationService.SaveConfiguration(config);
            
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
                
                // Update configuration (persist changes)
                var config = _configurationService.LoadConfiguration();
                config.VolunteerFilePath = _volunteerFilePath;
                _configurationService.SaveConfiguration(config);
                
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
}
