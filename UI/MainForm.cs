using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
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
        private Panel _sidebar = null!;
        private Panel _navTransform = null!;
        private Panel _navVolunteers = null!;
        private Label _navTransformLabel = null!;
        private Label _navVolunteersLabel = null!;
        private Panel _pageHeader = null!;
        private Label _lblPageTitle = null!;
        private Panel _contentPanel = null!;
        private Panel _innerPanel = null!;
        private RoundedPanel _transformPage = null!;
        private RoundedPanel _transformLeftCard = null!;

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
            this.Text = "Auser Gestione Trasporti v2.0.2";
            this.FormBorderStyle = FormBorderStyle.Sizable;
            this.WindowState = FormWindowState.Maximized;
            this.MinimumSize = new Size(700, 600);
            this.MaximizeBox = true;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = ThemeManager.ColorAppBackground;

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

            // Sidebar — fixed-width left navigation, styled after the inspiration
            // dashboard: app title + subtitle up top, pill-highlighted nav items below.
            _sidebar = new Panel
            {
                Dock = DockStyle.Left,
                Width = 280,
                BackColor = ThemeManager.ColorAppBackground,
                Padding = new Padding(16, 24, 16, 24)
            };
            _sidebar.Paint += (s, e) =>
            {
                using var pen = new Pen(ThemeManager.ColorBorderLight, 1);
                e.Graphics.DrawLine(pen, _sidebar.Width - 1, 0, _sidebar.Width - 1, _sidebar.Height);
            };
            this.Controls.Add(_sidebar);

            var lblAppTitle = new Label
            {
                Text = "Auser Gestione Trasporti",
                Location = new Point(16, 24),
                Size = new Size(248, 44),
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _sidebar.Controls.Add(lblAppTitle);

            var lblAppSubtitle = new Label
            {
                Text = "Area Volontari",
                Location = new Point(16, 64),
                Size = new Size(248, 20),
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorSecondary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            _sidebar.Controls.Add(lblAppSubtitle);

            _navTransform = CreateNavItem("⇄   Aggiungi Accompagnamenti", 100, out _navTransformLabel);
            _navVolunteers = CreateNavItem("👥   Volontari", 148, out _navVolunteersLabel);
            _sidebar.Controls.Add(_navTransform);
            _sidebar.Controls.Add(_navVolunteers);
            _navTransform.Click += (s, e) => SelectPage(true);
            _navVolunteers.Click += (s, e) => SelectPage(false);
            _navTransformLabel.Click += (s, e) => SelectPage(true);
            _navVolunteersLabel.Click += (s, e) => SelectPage(false);

            // Main container fills the remaining space to the right of the sidebar
            var mainContainer = new Panel { Dock = DockStyle.Fill, BackColor = ThemeManager.ColorAppBackground };
            this.Controls.Add(mainContainer);
            mainContainer.BringToFront();

            // PageHeader — slim top bar showing the current page's title
            _pageHeader = new Panel
            {
                Dock = DockStyle.Top,
                Height = 64,
                BackColor = ThemeManager.ColorAppBackground,
                Padding = new Padding(32, 0, 32, 0)
            };
            _pageHeader.Paint += (s, e) =>
            {
                using var pen = new Pen(ThemeManager.ColorBorderLight, 1);
                e.Graphics.DrawLine(pen, 0, _pageHeader.Height - 1, _pageHeader.Width, _pageHeader.Height - 1);
            };
            mainContainer.Controls.Add(_pageHeader);

            _lblPageTitle = new Label
            {
                Text = "Aggiungi Accompagnamenti",
                Location = new Point(32, 14),
                AutoSize = true,
                Font = new Font("Segoe UI", 15F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            _pageHeader.Controls.Add(_lblPageTitle);

            // ContentPanel — scrollable, fills remaining space below the page header
            _contentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                AutoScroll = true,
                BackColor = ThemeManager.ColorAppBackground
            };
            mainContainer.Controls.Add(_contentPanel);
            _contentPanel.BringToFront();

            // InnerPanel — centered, capped at a wide desktop-dashboard width
            _innerPanel = new Panel
            {
                Location = new Point(20, 20),
                Width = Math.Max(320, Math.Min(_contentPanel.ClientSize.Width - 32, 1080)),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                BackColor = ThemeManager.ColorAppBackground
            };
            _contentPanel.Controls.Add(_innerPanel);

            // Resize handler to recalculate InnerPanel position and width
            _contentPanel.Resize += ContentPanel_Resize;

            // TransformPage wraps the two transform cards so it can be shown/hidden as a "page"
            _transformPage = new RoundedPanel
            {
                Location = new Point(0, 0),
                Width = _innerPanel.Width,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent,
                HoverEffect = false,
                CornerRadius = 0,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink
            };
            _innerPanel.Controls.Add(_transformPage);

            // Single full-width card, matching the inspiration's "Importa Dati" panel,
            // with Elabora and Salva as a paired action row at the bottom.
            int leftWidth = _transformPage.Width;

            _transformLeftCard = new RoundedPanel
            {
                Location = new Point(0, 0),
                Size = new Size(leftWidth, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.White,
                CornerRadius = 16,
                AccentBarColor = ThemeManager.ColorPrimary
            };
            _transformPage.Controls.Add(_transformLeftCard);

            // --- Card: "Importa Dati (CSV/Excel)" ---
            var lblLeftTitle = new Label
            {
                Text = "Importa Dati (CSV/Excel)",
                Location = new Point(24, 20),
                Size = new Size(leftWidth - 48, 26),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            _transformLeftCard.Controls.Add(lblLeftTitle);

            var lblLeftSubtitle = new Label
            {
                Text = "Seleziona i file per avviare l'elaborazione dei turni.",
                Location = new Point(24, 48),
                Size = new Size(leftWidth - 48, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorSecondary
            };
            _transformLeftCard.Controls.Add(lblLeftSubtitle);

            btnSelectCSV = new ModernButton
            {
                Text = Properties.Resources.SelectCSVButton,
                Location = new Point(24, 86),
                Size = new Size(leftWidth - 48, 46),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlinePrimary(btnSelectCSV);
            btnSelectCSV.Click += BtnSelectCSV_Click;
            _transformLeftCard.Controls.Add(btnSelectCSV);

            lblCSVPath = new Label
            {
                Text = "",
                Location = new Point(24, 136),
                Size = new Size(leftWidth - 48, 22),
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorSecondary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _transformLeftCard.Controls.Add(lblCSVPath);

            btnSelectExcel = new ModernButton
            {
                Text = Properties.Resources.SelectExcelButton,
                Location = new Point(24, 168),
                Size = new Size(leftWidth - 48, 46),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlinePrimary(btnSelectExcel);
            btnSelectExcel.Click += BtnSelectExcel_Click;
            _transformLeftCard.Controls.Add(btnSelectExcel);

            lblExcelPath = new Label
            {
                Text = "",
                Location = new Point(24, 218),
                Size = new Size(leftWidth - 48, 22),
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorSecondary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                AutoEllipsis = true,
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft
            };
            _transformLeftCard.Controls.Add(lblExcelPath);

            var leftSeparator = new Panel
            {
                Location = new Point(0, 250),
                Size = new Size(leftWidth, 1),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = ThemeManager.ColorBorderLight
            };
            _transformLeftCard.Controls.Add(leftSeparator);

            // Elabora and Salva (export xlsx) share the action row, side by side.
            int actionGap = 16;
            int actionWidth = (leftWidth - 48 - actionGap) / 2;

            btnProcess = new ModernButton
            {
                Text = Properties.Resources.ProcessButton,
                Location = new Point(24, 262),
                Size = new Size(actionWidth, 48),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyPrimary(btnProcess);
            btnProcess.Click += BtnProcess_Click;
            _transformLeftCard.Controls.Add(btnProcess);

            btnDownload = new ModernButton
            {
                Text = "Salva",
                Location = new Point(24 + actionWidth + actionGap, 262),
                Size = new Size(actionWidth, 48),
                Enabled = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlinePrimary(btnDownload);
            btnDownload.Click += BtnDownload_Click;
            _transformLeftCard.Controls.Add(btnDownload);

            // Status label sits right under the action row so feedback is visible
            // exactly where the action was triggered.
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(24, 318),
                Size = new Size(leftWidth - 48, 40),
                Font = new Font("Segoe UI", 9F),
                AutoSize = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                BackColor = Color.Transparent
            };
            _transformLeftCard.Controls.Add(lblStatus);

            _transformLeftCard.Height = 374;

            SelectPage(true);
        }

        /// <summary>
        /// Creates a sidebar navigation item styled as a pill-highlighted row,
        /// matching the inspiration dashboard's side navigation.
        /// </summary>
        private Panel CreateNavItem(string text, int y, out Label textLabel)
        {
            var item = new Panel
            {
                Location = new Point(16, y),
                Size = new Size(248, 40),
                Cursor = Cursors.Hand,
                BackColor = Color.Transparent,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            item.Paint += (s, e) =>
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                bool active = (string)item.Tag == "active";
                if (active)
                {
                    using var path = RoundedRect(new Rectangle(0, 0, item.Width - 1, item.Height - 1), 10);
                    using var brush = new SolidBrush(ThemeManager.ColorSoft);
                    e.Graphics.FillPath(brush, path);
                }
            };
            var lbl = new Label
            {
                Text = text,
                Dock = DockStyle.Fill,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorSecondary,
                Padding = new Padding(12, 0, 0, 0),
                BackColor = Color.Transparent
            };
            item.Controls.Add(lbl);
            textLabel = lbl;
            return item;
        }

        private static GraphicsPath RoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private void SelectPage(bool transform)
        {
            _navTransform.Tag = transform ? "active" : "";
            _navVolunteers.Tag = !transform ? "active" : "";
            _navTransformLabel.ForeColor = transform ? ThemeManager.ColorPrimary : ThemeManager.ColorSecondary;
            _navVolunteersLabel.ForeColor = !transform ? ThemeManager.ColorPrimary : ThemeManager.ColorSecondary;
            _navTransform.Invalidate();
            _navVolunteers.Invalidate();

            _lblPageTitle.Text = transform ? "Aggiungi Accompagnamenti" : "Gestione Volontari";

            _transformPage.Visible = transform;
            if (_volunteerPanel != null)
                _volunteerPanel.Visible = !transform;

            _contentPanel.AutoScrollPosition = new Point(0, 0);
        }

        // Recalculate InnerPanel position and width on ContentPanel resize
        private void ContentPanel_Resize(object? sender, EventArgs e)
        {
            int available = _contentPanel.ClientSize.Width;
            // Capped so wide screens are put to use, while still shrinking gracefully
            // on narrower windows instead of stretching every control edge-to-edge.
            int maxWidth = 1080;
            int panelWidth = Math.Max(320, Math.Min(available - 32, maxWidth));
            int panelX = Math.Max(16, (available - panelWidth) / 2);
            _innerPanel.Location = new Point(panelX, 24);
            _innerPanel.Width = panelWidth;

            if (_transformPage != null) _transformPage.Width = panelWidth;
            if (_transformLeftCard != null)
            {
                _transformLeftCard.Width = panelWidth;

                int actionGap = 16;
                int actionWidth = (panelWidth - 48 - actionGap) / 2;

                foreach (Control c in _transformLeftCard.Controls)
                {
                    if (c == btnProcess || c == btnDownload) continue;
                    c.Width = panelWidth - 48;
                }
                if (btnProcess != null) btnProcess.Width = actionWidth;
                if (btnDownload != null) { btnDownload.Width = actionWidth; btnDownload.Left = 24 + actionWidth + actionGap; }
            }
            if (_volunteerPanel != null) _volunteerPanel.Width = panelWidth;
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

                // VolunteerPanel is its own "page", toggled via the sidebar navigation
                panel.Location = new Point(0, 0);
                panel.Size = new Size(_innerPanel.Width, 648);
                panel.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
                panel.Visible = false;
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

            public void ShowVolunteerErrorMessage(string message)
                => Panel?.ShowVolunteerErrorMessage(message);
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
                    lblStatus.ForeColor = Color.FromArgb(0x06, 0x85, 0x34);
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.FromArgb(0x06, 0x85, 0x34);
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
