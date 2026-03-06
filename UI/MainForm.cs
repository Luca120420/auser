using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.UI
{
    /// <summary>
    /// Main form for the Auser Excel Transformer application.
    /// Implements the IGUI interface to provide user interaction capabilities.
    /// Validates: Requirements 1.1, 1.4, 1.5, 7.3, 8.1, 8.2
    /// </summary>
    public partial class MainForm : Form, IGUI
    {
        private readonly IApplicationController _controller;
        private readonly VolunteerPanel? _volunteerPanel;
        
        // UI Controls
        private Button btnSelectCSV = null!;
        private Button btnSelectExcel = null!;
        private Button btnProcess = null!;
        private Button btnDownload = null!;
        private Label lblCSVFile = null!;
        private Label lblExcelFile = null!;
        private Label lblCSVPath = null!;
        private Label lblExcelPath = null!;
        private Label lblStatus = null!;
        
        /// <summary>
        /// Initializes a new instance of the MainForm class.
        /// </summary>
        /// <param name="controller">The application controller to handle business logic</param>
        public MainForm(IApplicationController controller)
        {
            _controller = controller ?? throw new ArgumentNullException(nameof(controller));
            InitializeComponent();
            InitializeCustomComponents();
            
            // Initialize volunteer feature with dependency injection (Requirements 1.6, 3.3)
            _volunteerPanel = InitializeVolunteerFeature();
        }
        
        /// <summary>
        /// Initializes the volunteer notification feature with all required dependencies.
        /// Creates service instances and the VolunteerPanel control.
        /// Validates: Requirements 1.6, 3.3, 10.1, 10.2, 10.3
        /// </summary>
        /// <returns>The initialized VolunteerPanel, or null if initialization fails</returns>
        private VolunteerPanel? InitializeVolunteerFeature()
        {
            try
            {
                // Create service instances (Requirement 10.2)
                var volunteerManager = new VolunteerManager();
                var emailService = new EmailService();
                var configurationService = new ConfigurationService(volunteerManager);
                var excelManager = new ExcelManager();
                
                // Create a simple wrapper that will hold the panel reference
                // This allows us to pass the UI to the controller before the panel is fully created
                VolunteerPanelWrapper wrapper = new VolunteerPanelWrapper();
                
                // Create VolunteerNotificationController with all dependencies (Requirement 10.2)
                var controller = new VolunteerNotificationController(
                    volunteerManager,
                    emailService,
                    configurationService,
                    excelManager,
                    wrapper);
                
                // Create the VolunteerPanel with the controller (Requirement 10.2)
                var panel = new VolunteerPanel(controller);
                wrapper.Panel = panel; // Set the actual panel in the wrapper
                
                // Add panel to form layout below existing transformation controls (Requirement 10.1)
                panel.Location = new Point(20, 275); // Reduced spacing from status label
                panel.Size = new Size(this.ClientSize.Width - 40, 620); // Increased height to accommodate all controls
                panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                this.Controls.Add(panel);
                
                // Form size is already set in InitializeCustomComponents to accommodate volunteer panel
                // No need to adjust size here
                
                // Trigger display of loaded configuration now that panel is ready (Requirement 10.3)
                // The controller already loaded the configuration in its constructor,
                // now we refresh the UI to display the loaded data.
                controller.RefreshUIDisplay();
                
                return panel;
            }
            catch (Exception ex)
            {
                // Log error and return null if initialization fails
                System.Diagnostics.Debug.WriteLine($"Failed to initialize volunteer feature: {ex.Message}");
                return null;
            }
        }
        
        /// <summary>
        /// Wrapper class to handle circular dependency between controller and panel.
        /// </summary>
        private class VolunteerPanelWrapper : IVolunteerUI
        {
            public VolunteerPanel? Panel { get; set; }
            
            public void DisplayVolunteerList(Dictionary<string, string> volunteers)
            {
                Panel?.DisplayVolunteerList(volunteers);
            }
            
            public void DisplayGmailCredentials(string email, string password)
            {
                Panel?.DisplayGmailCredentials(email, password);
            }
            
            public void DisplaySheetNames(List<string> sheetNames)
            {
                Panel?.DisplaySheetNames(sheetNames);
            }
            
            public void EnableSendEmailsButton(bool enabled)
            {
                Panel?.EnableSendEmailsButton(enabled);
            }
            
            public void ShowEmailProgress(string message)
            {
                Panel?.ShowEmailProgress(message);
            }
            
            public void ShowEmailSummary(int successCount, int failureCount)
            {
                Panel?.ShowEmailSummary(successCount, failureCount);
            }
            
            public bool ConfirmAction(string message)
            {
                return Panel?.ConfirmAction(message) ?? false;
            }
            
            public void ShowErrorMessage(string message)
            {
                Panel?.ShowErrorMessage(message);
            }
        }
        
        /// <summary>
        /// Initializes the form components programmatically.
        /// </summary>
        private void InitializeCustomComponents()
        {
            // Set form properties
            this.Text = Properties.Resources.ApplicationTitle;
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(600, 400);
            this.MaximizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.AutoScroll = true; // Enable vertical scrolling when content exceeds visible area
            
            // Set initial size to accommodate all controls including VolunteerPanel
            this.Size = new Size(850, 1000);
            
            // Initialize CSV file selection button
            btnSelectCSV = new Button
            {
                Text = Properties.Resources.SelectCSVButton,
                Location = new Point(20, 20),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnSelectCSV.Click += BtnSelectCSV_Click;
            this.Controls.Add(btnSelectCSV);
            
            // Initialize CSV file label
            lblCSVFile = new Label
            {
                Text = Properties.Resources.CSVFileLabel,
                Location = new Point(20, 70),
                Size = new Size(100, 20),
                Font = new Font("Segoe UI", 9F)
            };
            this.Controls.Add(lblCSVFile);
            
            // Initialize CSV path display label
            lblCSVPath = new Label
            {
                Text = "",
                Location = new Point(120, 70),
                Size = new Size(this.ClientSize.Width - 140, 20), // Dynamic width based on form size
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkBlue,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblCSVPath);
            
            // Initialize Excel file selection button
            btnSelectExcel = new Button
            {
                Text = Properties.Resources.SelectExcelButton,
                Location = new Point(20, 100),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnSelectExcel.Click += BtnSelectExcel_Click;
            this.Controls.Add(btnSelectExcel);
            
            // Initialize Excel file label
            lblExcelFile = new Label
            {
                Text = Properties.Resources.ExcelFileLabel,
                Location = new Point(20, 150),
                Size = new Size(100, 20),
                Font = new Font("Segoe UI", 9F)
            };
            this.Controls.Add(lblExcelFile);
            
            // Initialize Excel path display label
            lblExcelPath = new Label
            {
                Text = "",
                Location = new Point(120, 150),
                Size = new Size(this.ClientSize.Width - 140, 20), // Dynamic width based on form size
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkBlue,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblExcelPath);
            
            // Initialize process button (initially disabled)
            btnProcess = new Button
            {
                Text = Properties.Resources.ProcessButton,
                Location = new Point(20, 190),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnProcess.Click += BtnProcess_Click;
            this.Controls.Add(btnProcess);
            
            // Initialize download button (initially disabled)
            btnDownload = new Button
            {
                Text = Properties.Resources.DownloadButton,
                Location = new Point(240, 190),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            btnDownload.Click += BtnDownload_Click;
            this.Controls.Add(btnDownload);
            
            // Initialize status label for messages
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(20, 250),
                Size = new Size(this.ClientSize.Width - 40, 20), // Reduced height from 40 to 20
                Font = new Font("Segoe UI", 9F),
                AutoSize = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(lblStatus);
        }
        
        #region Event Handlers
        
        /// <summary>
        /// Handles the CSV file selection button click event.
        /// </summary>
        private void BtnSelectCSV_Click(object? sender, EventArgs e)
        {
            var filePath = SelectCSVFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                _controller.OnCSVFileSelected(filePath);
            }
        }
        
        /// <summary>
        /// Handles the Excel file selection button click event.
        /// </summary>
        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            var filePath = SelectExcelFile();
            if (!string.IsNullOrEmpty(filePath))
            {
                _controller.OnExcelFileSelected(filePath);
            }
        }
        
        /// <summary>
        /// Handles the process button click event.
        /// </summary>
        private void BtnProcess_Click(object? sender, EventArgs e)
        {
            _controller.OnProcessButtonClicked();
        }
        
        /// <summary>
        /// Handles the download button click event.
        /// </summary>
        private void BtnDownload_Click(object? sender, EventArgs e)
        {
            _controller.OnDownloadButtonClicked();
        }
        
        #endregion
        
        #region IGUI Implementation
        
        /// <summary>
        /// Shows the main application window.
        /// </summary>
        public void ShowWindow()
        {
            Application.Run(this);
        }
        
        /// <summary>
        /// Opens a file selection dialog for CSV files.
        /// Validates: Requirements 1.2, 8.5
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        public string? SelectCSVFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = Properties.Resources.SelectCSVDialogTitle;
                dialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }
            return null;
        }
        
        /// <summary>
        /// Opens a file selection dialog for Excel files.
        /// Validates: Requirements 1.3, 8.5
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        public string? SelectExcelFile()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = Properties.Resources.SelectExcelDialogTitle;
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }
            return null;
        }
        
        /// <summary>
        /// Displays the selected CSV file path in the GUI.
        /// Validates: Requirements 1.4
        /// </summary>
        /// <param name="path">The file path to display</param>
        public void DisplaySelectedCSVPath(string path)
        {
            if (lblCSVPath.InvokeRequired)
            {
                lblCSVPath.Invoke(new Action(() => lblCSVPath.Text = path));
            }
            else
            {
                lblCSVPath.Text = path;
            }
        }
        
        /// <summary>
        /// Displays the selected Excel file path in the GUI.
        /// Validates: Requirements 1.5
        /// </summary>
        /// <param name="path">The file path to display</param>
        public void DisplaySelectedExcelPath(string path)
        {
            if (lblExcelPath.InvokeRequired)
            {
                lblExcelPath.Invoke(new Action(() => lblExcelPath.Text = path));
            }
            else
            {
                lblExcelPath.Text = path;
            }
        }
        
        /// <summary>
        /// Enables or disables the process button.
        /// Validates: Requirements 7.3
        /// </summary>
        /// <param name="enabled">True to enable, false to disable</param>
        public void EnableProcessButton(bool enabled)
        {
            if (btnProcess.InvokeRequired)
            {
                btnProcess.Invoke(new Action(() => btnProcess.Enabled = enabled));
            }
            else
            {
                btnProcess.Enabled = enabled;
            }
        }
        
        /// <summary>
        /// Enables or disables the download button.
        /// Validates: Requirements 7.3
        /// </summary>
        /// <param name="enabled">True to enable, false to disable</param>
        public void EnableDownloadButton(bool enabled)
        {
            if (btnDownload.InvokeRequired)
            {
                btnDownload.Invoke(new Action(() => btnDownload.Enabled = enabled));
            }
            else
            {
                btnDownload.Enabled = enabled;
            }
        }
        
        /// <summary>
        /// Displays an error message to the user in Italian.
        /// Validates: Requirements 7.6, 8.3, 8.4
        /// </summary>
        /// <param name="message">The error message to display</param>
        public void ShowErrorMessage(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = Color.Red;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.Red;
            }
        }
        
        /// <summary>
        /// Displays a success message to the user in Italian.
        /// Validates: Requirements 7.6, 8.4
        /// </summary>
        /// <param name="message">The success message to display</param>
        public void ShowSuccessMessage(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = Color.Green;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.Green;
            }
        }
        
        /// <summary>
        /// Opens a save file dialog and returns the selected path.
        /// Validates: Requirements 7.4, 8.5
        /// </summary>
        /// <returns>The selected file path, or null if cancelled</returns>
        public string? GetSaveFilePath()
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = Properties.Resources.SaveFileDialogTitle;
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                dialog.DefaultExt = "xlsx";
                dialog.AddExtension = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }
            return null;
        }
        
        #endregion
    }
}
