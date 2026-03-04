using System;
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
        }
        
        /// <summary>
        /// Initializes the form components programmatically.
        /// </summary>
        private void InitializeCustomComponents()
        {
            // Set form properties
            this.Text = Properties.Resources.ApplicationTitle;
            this.Size = new Size(600, 400);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            
            // Initialize CSV file selection button
            btnSelectCSV = new Button
            {
                Text = Properties.Resources.SelectCSVButton,
                Location = new Point(20, 20),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F)
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
                Size = new Size(450, 20),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkBlue
            };
            this.Controls.Add(lblCSVPath);
            
            // Initialize Excel file selection button
            btnSelectExcel = new Button
            {
                Text = Properties.Resources.SelectExcelButton,
                Location = new Point(20, 100),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F)
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
                Size = new Size(450, 20),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.DarkBlue
            };
            this.Controls.Add(lblExcelPath);
            
            // Initialize process button (initially disabled)
            btnProcess = new Button
            {
                Text = Properties.Resources.ProcessButton,
                Location = new Point(20, 190),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Enabled = false
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
                Enabled = false
            };
            btnDownload.Click += BtnDownload_Click;
            this.Controls.Add(btnDownload);
            
            // Initialize status label for messages
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(20, 250),
                Size = new Size(550, 80),
                Font = new Font("Segoe UI", 9F),
                AutoSize = false
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
