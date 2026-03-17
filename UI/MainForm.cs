using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.UI.Controls;

namespace AuserExcelTransformer.UI
{
    /// <summary>
    /// Main form for the Auser Excel Transformer application.
    /// Implements the IGUI interface to provide user interaction capabilities.
    /// Validates: Requirements 2.1, 3.1-3.7, 4.1-4.8, 7.1-7.7, 9.1-9.6, 10.1-10.4
    /// </summary>
    public partial class MainForm : Form, IGUI
    {
        private readonly IApplicationController _controller;
        private readonly VolunteerPanel? _volunteerPanel;

        // Layout panels
        private HeaderPanel _headerPanel = null!;
        private Panel _contentPanel = null!;
        private Panel _innerPanel = null!;

        // Transform card controls
        private ModernButton btnSelectCSV = null!;
        private ModernButton btnSelectExcel = null!;
        private ModernButton btnProcess = null!;
        private ModernButton btnDownload = null!;
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

            // Initialize volunteer feature with dependency injection
            _volunteerPanel = InitializeVolunteerFeature();
        }

        /// <summary>
        /// Initializes the form components programmatically.
        /// Tasks 4.1-4.6: HeaderPanel, ContentPanel, InnerPanel, TransformCard, ModernButtons, MinimumSize, AutoEllipsis
        /// </summary>
        private void InitializeCustomComponents()
        {
            // Task 4.5: form properties
            this.Text = "Auser Gestione Trasporti";
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.MinimumSize = new Size(700, 600);
            this.MaximizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;

            // Load application icon
            try
            {
                var assembly = System.Reflection.Assembly.GetExecutingAssembly();
                using (var stream = assembly.GetManifestResourceStream("AuserExcelTransformer.Resources.app_icon.ico"))
                {
                    if (stream != null)
                        this.Icon = new Icon(stream);
                }
            }
            catch { /* continue without icon */ }

            // Task 4.1: HeaderPanel — fixed 80px, Anchor Top+Left+Right
            _headerPanel = new HeaderPanel
            {
                Location = new Point(0, 0),
                Width = this.ClientSize.Width,
                Height = 80,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            this.Controls.Add(_headerPanel);

            // Task 4.1: ContentPanel — scrollable, fills remaining space, Anchor Top+Bottom+Left+Right
            _contentPanel = new Panel
            {
                Location = new Point(0, 80),
                Size = new Size(this.ClientSize.Width, this.ClientSize.Height - 80),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                AutoScroll = true,
                BackColor = Color.White
            };
            this.Controls.Add(_contentPanel);

            // Task 4.2: InnerPanel — centered, max 900px, inside ContentPanel
            _innerPanel = new Panel
            {
                Location = new Point(20, 20),
                Width = Math.Min(_contentPanel.ClientSize.Width - 40, 900),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = Color.White
            };
            _contentPanel.Controls.Add(_innerPanel);

            // Task 4.2: Resize handler to recalculate InnerPanel position and width
            _contentPanel.Resize += ContentPanel_Resize;

            // Task 4.3: TransformCard — Panel with 1px #E0E0E0 border, 24px padding
            var transformCard = new Panel
            {
                Location = new Point(0, 0),
                Width = _innerPanel.Width,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                Padding = new Padding(24)
            };
            _innerPanel.Controls.Add(transformCard);

            // Row 1: [Seleziona CSV 180px] [lblCSVPath expanding]
            // Task 4.4: Secondary style for btnSelectCSV
            btnSelectCSV = new ModernButton
            {
                Text = Properties.Resources.SelectCSVButton,
                Location = new Point(24, 24),
                Size = new Size(180, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplySecondary(btnSelectCSV);
            btnSelectCSV.Click += BtnSelectCSV_Click;
            transformCard.Controls.Add(btnSelectCSV);

            // Task 4.6: AutoEllipsis = true on file path labels
            lblCSVPath = new Label
            {
                Text = "",
                Location = new Point(216, 32),
                Size = new Size(transformCard.Width - 240, 24),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(0x39, 0x39, 0x39),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            transformCard.Controls.Add(lblCSVPath);

            // Row 2: [Seleziona Excel 180px] [lblExcelPath expanding]
            // Task 4.4: Secondary style for btnSelectExcel
            btnSelectExcel = new ModernButton
            {
                Text = Properties.Resources.SelectExcelButton,
                Location = new Point(24, 80),
                Size = new Size(180, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplySecondary(btnSelectExcel);
            btnSelectExcel.Click += BtnSelectExcel_Click;
            transformCard.Controls.Add(btnSelectExcel);

            // Task 4.6: AutoEllipsis = true on file path labels
            lblExcelPath = new Label
            {
                Text = "",
                Location = new Point(216, 88),
                Size = new Size(transformCard.Width - 240, 24),
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.FromArgb(0x39, 0x39, 0x39),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            transformCard.Controls.Add(lblExcelPath);

            // Row 3: [Elabora] [Scarica] — Task 4.4: Primary style
            btnProcess = new ModernButton
            {
                Text = Properties.Resources.ProcessButton,
                Location = new Point(24, 136),
                Size = new Size(160, 40),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyPrimary(btnProcess);
            btnProcess.Click += BtnProcess_Click;
            transformCard.Controls.Add(btnProcess);

            btnDownload = new ModernButton
            {
                Text = Properties.Resources.DownloadButton,
                Location = new Point(196, 136),
                Size = new Size(160, 40),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyPrimary(btnDownload);
            btnDownload.Click += BtnDownload_Click;
            transformCard.Controls.Add(btnDownload);

            // Row 4: status label
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(24, 192),
                Size = new Size(transformCard.Width - 48, 24),
                Font = new Font("Segoe UI", 9F),
                AutoSize = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            transformCard.Controls.Add(lblStatus);

            // Fix TransformCard height to fit its content
            transformCard.Height = 220;
        }

        // Task 4.2: Recalculate InnerPanel position and width on ContentPanel resize
        private void ContentPanel_Resize(object? sender, EventArgs e)
        {
            int available = _contentPanel.ClientSize.Width;
            int panelWidth = Math.Min(available - 40, 900);
            int panelX = Math.Max(20, (available - panelWidth) / 2);
            _innerPanel.Location = new Point(panelX, 20);
            _innerPanel.Width = panelWidth;

            // Keep TransformCard width in sync with InnerPanel
            if (_innerPanel.Controls.Count > 0 && _innerPanel.Controls[0] is Panel card)
            {
                card.Width = panelWidth;
                // Resize path labels inside the card
                foreach (Control c in card.Controls)
                {
                    if (c == lblCSVPath || c == lblExcelPath)
                        c.Width = panelWidth - 240;
                    if (c == lblStatus)
                        c.Width = panelWidth - 48;
                }
            }
        }

        /// <summary>
        /// Initializes the volunteer notification feature with all required dependencies.
        /// Adds VolunteerPanel to InnerPanel below the TransformCard.
        /// </summary>
        private VolunteerPanel? InitializeVolunteerFeature()
        {
            try
            {
                var volunteerManager = new VolunteerManager();
                var emailService = new EmailService();
                var configurationService = new ConfigurationService(volunteerManager);
                var excelManager = new ExcelManager();

                VolunteerPanelWrapper wrapper = new VolunteerPanelWrapper();

                var controller = new VolunteerNotificationController(
                    volunteerManager,
                    emailService,
                    configurationService,
                    excelManager,
                    wrapper);

                var panel = new VolunteerPanel(controller);
                wrapper.Panel = panel;

                // Add VolunteerPanel to InnerPanel below TransformCard (y=220+8=228)
                panel.Location = new Point(0, 228);
                panel.Size = new Size(_innerPanel.Width, 620);
                panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                _innerPanel.Controls.Add(panel);

                controller.RefreshUIDisplay();

                return panel;
            }
            catch (Exception ex)
            {
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
                => Panel?.DisplayVolunteerList(volunteers);

            public void DisplayGmailCredentials(string email, string password)
                => Panel?.DisplayGmailCredentials(email, password);

            public void DisplaySheetNames(List<string> sheetNames)
                => Panel?.DisplaySheetNames(sheetNames);

            public void EnableSendEmailsButton(bool enabled)
                => Panel?.EnableSendEmailsButton(enabled);

            public void ShowEmailProgress(string message)
                => Panel?.ShowEmailProgress(message);

            public void ShowEmailSummary(int successCount, int failureCount)
                => Panel?.ShowEmailSummary(successCount, failureCount);

            public bool ConfirmAction(string message)
                => Panel?.ConfirmAction(message) ?? false;

            public void ShowErrorMessage(string message)
                => Panel?.ShowErrorMessage(message);
        }

        #region Event Handlers

        private void BtnSelectCSV_Click(object? sender, EventArgs e)
        {
            var filePath = SelectCSVFile();
            if (!string.IsNullOrEmpty(filePath))
                _controller.OnCSVFileSelected(filePath);
        }

        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            var filePath = SelectExcelFile();
            if (!string.IsNullOrEmpty(filePath))
                _controller.OnExcelFileSelected(filePath);
        }

        private void BtnProcess_Click(object? sender, EventArgs e)
            => _controller.OnProcessButtonClicked();

        private void BtnDownload_Click(object? sender, EventArgs e)
            => _controller.OnDownloadButtonClicked();

        #endregion

        #region IGUI Implementation

        public void ShowWindow()
            => Application.Run(this);

        public string? SelectCSVFile()
        {
            using var dialog = new OpenFileDialog
            {
                Title = Properties.Resources.SelectCSVDialogTitle,
                Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };
            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        }

        public string? SelectExcelFile()
        {
            using var dialog = new OpenFileDialog
            {
                Title = Properties.Resources.SelectExcelDialogTitle,
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true
            };
            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        }

        public void DisplaySelectedCSVPath(string path)
        {
            if (lblCSVPath.InvokeRequired)
                lblCSVPath.Invoke(new Action(() => lblCSVPath.Text = path));
            else
                lblCSVPath.Text = path;
        }

        public void DisplaySelectedExcelPath(string path)
        {
            if (lblExcelPath.InvokeRequired)
                lblExcelPath.Invoke(new Action(() => lblExcelPath.Text = path));
            else
                lblExcelPath.Text = path;
        }

        public void EnableProcessButton(bool enabled)
        {
            if (btnProcess.InvokeRequired)
                btnProcess.Invoke(new Action(() => btnProcess.Enabled = enabled));
            else
                btnProcess.Enabled = enabled;
        }

        public void EnableDownloadButton(bool enabled)
        {
            if (btnDownload.InvokeRequired)
                btnDownload.Invoke(new Action(() => btnDownload.Enabled = enabled));
            else
                btnDownload.Enabled = enabled;
        }

        public void ShowErrorMessage(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = Color.FromArgb(0xD3, 0x2F, 0x2F);
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.FromArgb(0xD3, 0x2F, 0x2F);
            }
        }

        public void ShowSuccessMessage(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = Color.FromArgb(0x00, 0x92, 0x46);
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.FromArgb(0x00, 0x92, 0x46);
            }
        }

        public string? GetSaveFilePath()
        {
            using var dialog = new SaveFileDialog
            {
                Title = Properties.Resources.SaveFileDialogTitle,
                Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                DefaultExt = "xlsx",
                AddExtension = true
            };
            return dialog.ShowDialog() == DialogResult.OK ? dialog.FileName : null;
        }

        #endregion
    }
}
