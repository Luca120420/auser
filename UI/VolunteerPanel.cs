using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.UI
{
    /// <summary>
    /// User control for volunteer email notification management.
    /// Provides UI for managing volunteer contacts, Gmail credentials, Excel selection, and email sending.
    /// Implements IVolunteerUI interface.
    /// Validates: Requirements 1.1, 2.1, 2.3, 2.4, 3.1, 5.1, 8.1, 8.2, 8.6, 8.9
    /// </summary>
    public partial class VolunteerPanel : UserControl, IVolunteerUI
    {
        private readonly IVolunteerNotificationController _controller;
        
        // Volunteer Contacts Section (left column, main card)
        private Controls.RoundedPanel grpVolunteerContacts = null!;
        private ListView lstVolunteers = null!;
        private Controls.ModernButton btnAddVolunteers = null!;
        private Controls.ModernButton btnAddContact = null!;
        private Controls.ModernButton btnDeleteAll = null!;
        private Controls.ModernButton btnImportVolunteers = null!;
        private Controls.ModernButton btnExportVolunteers = null!;
        private Label lblVolunteerError = null!;
        
        // Gmail Credentials Section (right column)
        private Controls.RoundedPanel grpGmailCredentials = null!;
        private Label lblGmailEmail = null!;
        private Controls.ModernTextBox txtGmailEmail = null!;
        private Label lblGmailPassword = null!;
        private Controls.ModernTextBox txtGmailPassword = null!;
        private Controls.ModernButton btnClearCredentials = null!;
        private Controls.ModernButton btnSaveCredentials = null!;
        private Controls.ModernButton btnImportCredentials = null!;
        private Controls.ModernButton btnExportCredentials = null!;
        
        // Excel Selection Section (right column)
        private Controls.RoundedPanel grpExcelSelection = null!;
        private Controls.ModernButton btnSelectExcel = null!;
        private Label lblSheet = null!;
        private ComboBox cmbSheets = null!;
        
        // Email Sending Section (right column)
        private Controls.RoundedPanel grpEmailSending = null!;
        private Controls.ModernButton btnSendEmails = null!;
        private ProgressBar progressBar = null!;
        private Label lblStatus = null!;

        // Two-column layout metrics, shared across section builders and OnResize
        private int _leftColWidth;
        private int _rightColWidth;
        private const int ColumnGap = 24;
        
        /// <summary>
        /// Initializes a new instance of the VolunteerPanel class.
        /// </summary>
        /// <param name="controller">The volunteer notification controller</param>
        public VolunteerPanel(IVolunteerNotificationController controller)
        {
            _controller = controller ?? throw new ArgumentNullException(nameof(controller));
            InitializeComponent();
            InitializeCustomComponents();
        }
        
        /// <summary>
        /// Initializes all UI components programmatically.
        /// </summary>
        private void InitializeCustomComponents()
        {
            this.Size = new Size(1080, 648);
            this.AutoScroll = false; // Disable panel scrolling - let MainForm handle scrolling

            _leftColWidth = (int)((this.Width - ColumnGap) * 0.62);
            _rightColWidth = this.Width - ColumnGap - _leftColWidth;

            InitializeVolunteerContactsSection();
            InitializeGmailCredentialsSection();
            InitializeExcelSelectionSection();
            InitializeEmailSendingSection();

            this.Resize += VolunteerPanel_Resize;
        }

        /// <summary>
        /// Keeps the two-column card layout in sync when the panel is resized
        /// (e.g. when the window is resized and MainForm stretches this panel).
        /// </summary>
        private void VolunteerPanel_Resize(object? sender, EventArgs e)
        {
            _leftColWidth = (int)((this.Width - ColumnGap) * 0.62);
            _rightColWidth = this.Width - ColumnGap - _leftColWidth;
            int rightX = _leftColWidth + ColumnGap;

            if (grpVolunteerContacts != null)
            {
                grpVolunteerContacts.Width = _leftColWidth;
                foreach (Control c in grpVolunteerContacts.Controls)
                {
                    if (c == btnAddVolunteers || c == btnAddContact || c == btnDeleteAll) continue;
                    c.Width = grpVolunteerContacts.Width - 48;
                }
                int contactsBtnGap = 16;
                int contactsBtnWidth = (_leftColWidth - 48 - contactsBtnGap * 2) / 3;
                if (btnAddVolunteers != null) btnAddVolunteers.Width = contactsBtnWidth;
                if (btnAddContact != null) { btnAddContact.Width = contactsBtnWidth; btnAddContact.Left = 24 + contactsBtnWidth + contactsBtnGap; }
                if (btnDeleteAll != null) { btnDeleteAll.Width = contactsBtnWidth; btnDeleteAll.Left = 24 + (contactsBtnWidth + contactsBtnGap) * 2; }

                int importExportGap = 16;
                int importExportWidth = (_leftColWidth - 48 - importExportGap) / 2;
                if (btnImportVolunteers != null) btnImportVolunteers.Width = importExportWidth;
                if (btnExportVolunteers != null) { btnExportVolunteers.Width = importExportWidth; btnExportVolunteers.Left = 24 + importExportWidth + importExportGap; }
            }

            foreach (var card in new[] { grpGmailCredentials, grpExcelSelection, grpEmailSending })
            {
                if (card == null) continue;
                card.Left = rightX;
                card.Width = _rightColWidth;
            }
        }
        
        /// <summary>
        /// Initializes the volunteer contacts section.
        /// Validates: Requirements 1.1, 8.1, 8.2, 8.6, 8.9
        /// </summary>
        private void InitializeVolunteerContactsSection()
        {
            // Left column: main "Elenco Volontari" card, matching the inspiration's
            // primary content card with title + subtitle header and accent bar.
            int cardHeight = 648;
            grpVolunteerContacts = new Controls.RoundedPanel
            {
                Location = new Point(0, 0),
                Size = new Size(_leftColWidth, cardHeight),
                Anchor = AnchorStyles.Top | AnchorStyles.Left,
                BackColor = Color.White,
                CornerRadius = 16,
                AccentBarColor = ThemeManager.ColorPrimary
            };

            var lblTitle = new Label
            {
                Text = "Elenco Volontari",
                Location = new Point(24, 20),
                Size = new Size(grpVolunteerContacts.Width - 48, 26),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            grpVolunteerContacts.Controls.Add(lblTitle);

            var lblSubtitle = new Label
            {
                Text = "Gestisci le informazioni di contatto dei volontari.",
                Location = new Point(24, 48),
                Size = new Size(grpVolunteerContacts.Width - 48, 20),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorSecondary
            };
            grpVolunteerContacts.Controls.Add(lblSubtitle);

            // Error label for volunteer-list validation/import/delete errors (e.g. "Il
            // cognome non può essere vuoto.", "L'indirizzo email non è valido."). Shown
            // right under the section subtitle so it's clearly scoped to this card,
            // rather than the generic status label under "Invio Email". Empty by default.
            lblVolunteerError = new Label
            {
                Text = string.Empty,
                Location = new Point(24, 68),
                Size = new Size(grpVolunteerContacts.Width - 48, 18),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = ThemeManager.ColorError,
                AutoEllipsis = true
            };
            grpVolunteerContacts.Controls.Add(lblVolunteerError);

            // Three actions share one row, mobile/desktop-friendly, per the inspiration's
            // "Aggiungi Volontario" / "Elimina Tutto" action row above the list.
            int contactsBtnGap = 16;
            int contactsBtnWidth = (grpVolunteerContacts.Width - 48 - contactsBtnGap * 2) / 3;

            btnAddVolunteers = new Controls.ModernButton
            {
                Text = Properties.Resources.AddVolunteersButton,
                Location = new Point(24, 96),
                Size = new Size(contactsBtnWidth, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyPrimary(btnAddVolunteers);
            btnAddVolunteers.Click += BtnAddVolunteers_Click;
            grpVolunteerContacts.Controls.Add(btnAddVolunteers);

            btnAddContact = new Controls.ModernButton
            {
                Text = Properties.Resources.AddContactButton,
                Location = new Point(24 + contactsBtnWidth + contactsBtnGap, 96),
                Size = new Size(contactsBtnWidth, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyOutlinePrimary(btnAddContact);
            btnAddContact.Click += BtnAddContact_Click;
            grpVolunteerContacts.Controls.Add(btnAddContact);

            btnDeleteAll = new Controls.ModernButton
            {
                Text = Properties.Resources.DeleteAllButton,
                Location = new Point(24 + (contactsBtnWidth + contactsBtnGap) * 2, 96),
                Size = new Size(contactsBtnWidth, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlineNeutral(btnDeleteAll);
            btnDeleteAll.Click += BtnDeleteAll_Click;
            grpVolunteerContacts.Controls.Add(btnDeleteAll);

            // Second action row: "Importa" / "Esporta" for the volunteer list, mirroring
            // the same Importa/Esporta pairing used in the Credenziali Gmail card, so the
            // whole list (not just one contact at a time) can be backed up or restored.
            int importExportGap = 16;
            int importExportWidth = (grpVolunteerContacts.Width - 48 - importExportGap) / 2;

            btnImportVolunteers = new Controls.ModernButton
            {
                Text = "Importa",
                Location = new Point(24, 144),
                Size = new Size(importExportWidth, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyOutlinePrimary(btnImportVolunteers);
            btnImportVolunteers.Click += BtnImportVolunteers_Click;
            grpVolunteerContacts.Controls.Add(btnImportVolunteers);

            btnExportVolunteers = new Controls.ModernButton
            {
                Text = "Esporta",
                Location = new Point(24 + importExportWidth + importExportGap, 144),
                Size = new Size(importExportWidth, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlinePrimary(btnExportVolunteers);
            btnExportVolunteers.Click += BtnExportVolunteers_Click;
            grpVolunteerContacts.Controls.Add(btnExportVolunteers);

            // ListView for volunteer contacts, styled with avatar-initial circles per row
            lstVolunteers = new ListView
            {
                Location = new Point(24, 196),
                Size = new Size(grpVolunteerContacts.Width - 48, cardHeight - 196 - 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right,
                View = View.Details,
                FullRowSelect = true,
                GridLines = false,
                Font = new Font("Segoe UI", 9F),
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            StyleListView(lstVolunteers);

            // Add columns
            lstVolunteers.Columns.Add(Properties.Resources.VolunteerListColumnSurname, 300);
            lstVolunteers.Columns.Add(Properties.Resources.VolunteerListColumnEmail, 400);

            // Handle mouse click for delete action
            lstVolunteers.MouseClick += LstVolunteers_MouseClick;

            grpVolunteerContacts.Controls.Add(lstVolunteers);

            this.Controls.Add(grpVolunteerContacts);
        }

        /// <summary>
        /// Applies palette-consistent styling and a subtle hover highlight to a ListView.
        /// </summary>
        private static void StyleListView(ListView lv)
        {
            // A dummy ImageList with a tall ImageSize forces ListView to use taller
            // rows in Details view, giving room for the avatar-initial circles.
            var rowHeightImages = new ImageList { ImageSize = new Size(1, 36) };
            rowHeightImages.Images.Add(new Bitmap(1, 36));
            lv.SmallImageList = rowHeightImages;

            lv.OwnerDraw = true;
            lv.DrawColumnHeader += (s, e) =>
            {
                e.Graphics.FillRectangle(new SolidBrush(ThemeManager.ColorAppBackground), e.Bounds);
                TextRenderer.DrawText(e.Graphics, e.Header!.Text, new Font("Segoe UI", 9F, FontStyle.Bold),
                    e.Bounds, ThemeManager.ColorPrimary, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
            };
            int hoveredIndex = -1;
            lv.MouseMove += (s, e) =>
            {
                var item = lv.GetItemAt(e.X, e.Y);
                int idx = item?.Index ?? -1;
                if (idx != hoveredIndex) { hoveredIndex = idx; lv.Invalidate(); }
            };
            lv.MouseLeave += (s, e) => { hoveredIndex = -1; lv.Invalidate(); };
            lv.DrawItem += (s, e) =>
            {
                bool selected = e.Item.Selected;
                bool hovered = e.ItemIndex == hoveredIndex;
                var bg = selected ? ThemeManager.ColorSoft : (hovered ? ThemeManager.ColorAppBackground : Color.White);
                using var brush = new SolidBrush(bg);
                e.Graphics.FillRectangle(brush, e.Bounds);
            };
            lv.DrawSubItem += (s, e) =>
            {
                bool selected = e.Item!.Selected;
                bool hovered = e.ItemIndex == hoveredIndex;
                var bg = selected ? ThemeManager.ColorSoft : (hovered ? ThemeManager.ColorAppBackground : Color.White);
                using var brush = new SolidBrush(bg);
                e.Graphics.FillRectangle(brush, e.Bounds);

                if (e.ColumnIndex == 0)
                {
                    // Avatar-initial circle + name, matching the inspiration's list rows
                    string name = e.SubItem!.Text ?? "";
                    string initials = name.Length >= 2 ? name.Substring(0, 2).ToUpperInvariant()
                        : (name.Length == 1 ? name.ToUpperInvariant() : "?");

                    int diameter = 28;
                    int cx = e.Bounds.Left + 8;
                    int cy = e.Bounds.Top + (e.Bounds.Height - diameter) / 2;
                    var circleRect = new Rectangle(cx, cy, diameter, diameter);

                    e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                    using (var circleBrush = new SolidBrush(ThemeManager.ColorSoft))
                        e.Graphics.FillEllipse(circleBrush, circleRect);
                    TextRenderer.DrawText(e.Graphics, initials, new Font("Segoe UI", 8F, FontStyle.Bold),
                        circleRect, ThemeManager.ColorPrimary, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter);

                    var textBounds = new Rectangle(cx + diameter + 10, e.Bounds.Top, e.Bounds.Width - diameter - 18, e.Bounds.Height);
                    TextRenderer.DrawText(e.Graphics, name, new Font(lv.Font, FontStyle.Bold), textBounds, ThemeManager.ColorPrimary,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.LeftAndRightPadding);
                }
                else
                {
                    TextRenderer.DrawText(e.Graphics, e.SubItem!.Text, lv.Font, e.Bounds, ThemeManager.ColorPrimary,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.LeftAndRightPadding);
                }
            };
        }
        
        /// <summary>
        /// Initializes the Gmail credentials section.
        /// Validates: Requirements 3.1
        /// </summary>
        private void InitializeGmailCredentialsSection()
        {
            // Right column, card 1 of 3: "Credenziali Gmail" — mirrors the inspiration's
            // security card with icon header, info box, floating-style labeled inputs,
            // and a primary action followed by two secondary actions.
            int rightX = _leftColWidth + ColumnGap;
            grpGmailCredentials = new Controls.RoundedPanel
            {
                Location = new Point(rightX, 0),
                Size = new Size(_rightColWidth, 300),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.White,
                CornerRadius = 16,
                AccentBarColor = ThemeManager.ColorSecondaryAlt
            };

            var lblTitle = new Label
            {
                Text = "Credenziali Gmail",
                Location = new Point(24, 20),
                Size = new Size(grpGmailCredentials.Width - 48, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            grpGmailCredentials.Controls.Add(lblTitle);

            var lblSubtitle = new Label
            {
                Text = "Impostazioni di autenticazione del sistema.",
                Location = new Point(24, 46),
                Size = new Size(grpGmailCredentials.Width - 48, 18),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = ThemeManager.ColorSecondary
            };
            grpGmailCredentials.Controls.Add(lblSubtitle);

            // Gmail email label
            lblGmailEmail = new Label
            {
                Text = "Indirizzo Email del Servizio",
                Location = new Point(24, 76),
                Size = new Size(grpGmailCredentials.Width - 48, 18),
                Font = new Font("Segoe UI", 8F),
                ForeColor = ThemeManager.ColorSecondary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            grpGmailCredentials.Controls.Add(lblGmailEmail);

            // Gmail email textbox
            txtGmailEmail = new Controls.ModernTextBox
            {
                Location = new Point(24, 96),
                Size = new Size(grpGmailCredentials.Width - 48, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9F),
                ForeColor = Color.Black
            };
            txtGmailEmail.TextChanged += TxtGmailEmail_TextChanged;
            grpGmailCredentials.Controls.Add(txtGmailEmail);

            // Gmail password label
            lblGmailPassword = new Label
            {
                Text = "Password per l'app",
                Location = new Point(24, 134),
                Size = new Size(grpGmailCredentials.Width - 48, 18),
                Font = new Font("Segoe UI", 8F),
                ForeColor = ThemeManager.ColorSecondary,
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            grpGmailCredentials.Controls.Add(lblGmailPassword);

            // Gmail password textbox
            txtGmailPassword = new Controls.ModernTextBox
            {
                Location = new Point(24, 154),
                Size = new Size(grpGmailCredentials.Width - 48, 30),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9F),
                UseSystemPasswordChar = true,
                ForeColor = Color.Black
            };
            txtGmailPassword.TextChanged += TxtGmailPassword_TextChanged;
            grpGmailCredentials.Controls.Add(txtGmailPassword);

            // Primary action full width, then two secondary actions side by side —
            // matching the inspiration's "Salva Credenziali" / "Importa" / "Cancella" stack.
            btnSaveCredentials = new Controls.ModernButton
            {
                Text = "Salva Credenziali",
                Location = new Point(24, 198),
                Size = new Size(grpGmailCredentials.Width - 48, 42),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            ThemeManager.ApplyPrimary(btnSaveCredentials);
            btnSaveCredentials.Click += BtnSaveCredentials_Click;
            grpGmailCredentials.Controls.Add(btnSaveCredentials);

            // Three secondary actions share one row: Importa / Esporta / Cancella.
            int credBtnGap = 12;
            int credBtnWidth = (grpGmailCredentials.Width - 48 - credBtnGap * 2) / 3;

            btnImportCredentials = new Controls.ModernButton
            {
                Text = "Importa",
                Location = new Point(24, 248),
                Size = new Size(credBtnWidth, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyOutlinePrimary(btnImportCredentials);
            btnImportCredentials.Click += BtnImportCredentials_Click;
            grpGmailCredentials.Controls.Add(btnImportCredentials);

            btnExportCredentials = new Controls.ModernButton
            {
                Text = "Esporta",
                Location = new Point(24 + credBtnWidth + credBtnGap, 248),
                Size = new Size(credBtnWidth, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            ThemeManager.ApplyOutlinePrimary(btnExportCredentials);
            btnExportCredentials.Click += BtnExportCredentials_Click;
            grpGmailCredentials.Controls.Add(btnExportCredentials);

            btnClearCredentials = new Controls.ModernButton
            {
                Text = "Cancella",
                Location = new Point(24 + (credBtnWidth + credBtnGap) * 2, 248),
                Size = new Size(credBtnWidth, 38),
                Anchor = AnchorStyles.Top | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlineNeutral(btnClearCredentials);
            btnClearCredentials.Click += BtnClearCredentials_Click;
            grpGmailCredentials.Controls.Add(btnClearCredentials);

            grpGmailCredentials.Height = 300;

            this.Controls.Add(grpGmailCredentials);
        }
        
        /// <summary>
        /// Initializes the Excel selection section.
        /// Validates: Requirements 2.1, 2.3, 2.4
        /// </summary>
        private void InitializeExcelSelectionSection()
        {
            // Right column, card 2 of 3: "Selezione File Excel" — stacked layout since
            // the narrower right column doesn't have room for a single wide row.
            int rightX = _leftColWidth + ColumnGap;
            int cardY = 324; // below the Gmail card (300h + 24 gap)
            grpExcelSelection = new Controls.RoundedPanel
            {
                Location = new Point(rightX, cardY),
                Size = new Size(_rightColWidth, 150),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.White,
                CornerRadius = 16,
                AccentBarColor = ThemeManager.ColorPrimary
            };

            var lblTitle = new Label
            {
                Text = "Selezione File Excel",
                Location = new Point(24, 20),
                Size = new Size(grpExcelSelection.Width - 48, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            grpExcelSelection.Controls.Add(lblTitle);

            btnSelectExcel = new Controls.ModernButton
            {
                Text = "Seleziona File Excel",
                Location = new Point(24, 54),
                Size = new Size(grpExcelSelection.Width - 48, 40),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };
            ThemeManager.ApplyOutlinePrimary(btnSelectExcel);
            btnSelectExcel.Click += BtnSelectExcel_Click;
            grpExcelSelection.Controls.Add(btnSelectExcel);

            // Sheet label
            lblSheet = new Label
            {
                Text = "Foglio:",
                Location = new Point(24, 104),
                Size = new Size(56, 30),
                Font = new Font("Segoe UI", 9F),
                ForeColor = ThemeManager.ColorPrimary,
                TextAlign = ContentAlignment.MiddleLeft,
                Anchor = AnchorStyles.Top | AnchorStyles.Left
            };
            grpExcelSelection.Controls.Add(lblSheet);

            // Sheet combobox
            cmbSheets = new ComboBox
            {
                Location = new Point(80, 104),
                Size = new Size(grpExcelSelection.Width - 104, 26),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 9F),
                DropDownStyle = ComboBoxStyle.DropDownList,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Color.Black
            };
            cmbSheets.SelectedIndexChanged += CmbSheets_SelectedIndexChanged;
            grpExcelSelection.Controls.Add(cmbSheets);

            this.Controls.Add(grpExcelSelection);
        }
        
        /// <summary>
        /// Initializes the email sending section.
        /// Validates: Requirements 5.1
        /// </summary>
        private void InitializeEmailSendingSection()
        {
            // Right column, card 3 of 3: "Invio Email" — below the Excel card
            // (150h + 24 gap after the 324 offset of card 2).
            int rightX = _leftColWidth + ColumnGap;
            int cardY = 324 + 150 + 24;
            grpEmailSending = new Controls.RoundedPanel
            {
                Location = new Point(rightX, cardY),
                Size = new Size(_rightColWidth, 190),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                BackColor = Color.White,
                CornerRadius = 16,
                AccentBarColor = ThemeManager.ColorSecondaryAlt
            };

            var lblTitle = new Label
            {
                Text = "Invio Email",
                Location = new Point(24, 20),
                Size = new Size(grpEmailSending.Width - 48, 24),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 12F, FontStyle.Bold),
                ForeColor = ThemeManager.ColorPrimary
            };
            grpEmailSending.Controls.Add(lblTitle);

            btnSendEmails = new Controls.ModernButton
            {
                Text = Properties.Resources.SendEmailsButton,
                Location = new Point(24, 54),
                Size = new Size(grpEmailSending.Width - 48, 42),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Enabled = false
            };
            ThemeManager.ApplyPrimary(btnSendEmails);
            btnSendEmails.Click += BtnSendEmails_Click;
            grpEmailSending.Controls.Add(btnSendEmails);

            // Progress bar. Not shown in the UI (never added to the visible control
            // tree below) since it was just a static, always-empty gray track that
            // ate up space above the status text without ever indicating progress.
            // The field is still created so send-progress logic (and existing tests
            // that look it up by field name) keep working if it's wired up later.
            progressBar = new ProgressBar
            {
                Location = new Point(24, 106),
                Size = new Size(grpEmailSending.Width - 48, 16),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Style = ProgressBarStyle.Continuous,
                Visible = false
            };

            // Status label. AutoSize is explicitly off and the label now starts where
            // the progress bar used to sit and fills the rest of the card, so long
            // messages (e.g. the "Invio completato..." summary) have much more room
            // to wrap instead of being squeezed into a thin strip.
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(24, 106),
                Size = new Size(grpEmailSending.Width - 48, 74),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right,
                Font = new Font("Segoe UI", 8.5F),
                AutoSize = false,
                AutoEllipsis = false
            };
            grpEmailSending.Controls.Add(lblStatus);

            this.Controls.Add(grpEmailSending);
        }
        
        #region Event Handlers
        
        /// <summary>
        /// Handles the Add Volunteers button click event.
        /// Opens file dialog to select volunteer JSON file.
        /// </summary>
        private void BtnAddVolunteers_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Seleziona file volontari";
                dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _controller.OnVolunteerFileSelected(dialog.FileName);
                }
            }
        }
        
        /// <summary>
        /// Handles the Importa (volunteer list) button click event.
        /// Behaves like "Aggiungi Volontari": loads a volunteer JSON file and merges
        /// it into the current list, but is offered here under the same "Importa" /
        /// "Esporta" naming used for Gmail credentials, for a consistent workflow.
        /// </summary>
        private void BtnImportVolunteers_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Importa Elenco Volontari";
                dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _controller.OnVolunteerFileSelected(dialog.FileName);
                }
            }
        }

        /// <summary>
        /// Handles the Esporta (volunteer list) button click event.
        /// Writes the current volunteer list to a JSON file using the same
        /// {"associates": { "Cognome": "email", ... }} format read by "Importa"
        /// and "Aggiungi Volontari", so the exported file can be re-imported.
        /// </summary>
        private void BtnExportVolunteers_Click(object? sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "Esporta Elenco Volontari";
                dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                dialog.FileName = "volontari-auser.json";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var associates = _controller.GetVolunteers();
                        var payload = new { associates };
                        var json = System.Text.Json.JsonSerializer.Serialize(payload, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                        System.IO.File.WriteAllText(dialog.FileName, json);
                        MessageBox.Show("Elenco volontari esportato.", "Esporta Elenco Volontari", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Errore nell'esportazione: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Handles mouse click on volunteer list for context menu.
        /// </summary>
        private void LstVolunteers_MouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var item = lstVolunteers.GetItemAt(e.X, e.Y);
                if (item != null)
                {
                    // Show context menu with edit and delete options
                    var contextMenu = new ContextMenuStrip();

                    var editItem = new ToolStripMenuItem("Modifica");
                    editItem.Click += (s, args) =>
                    {
                        var surname = item.Text;
                        var email = item.SubItems.Count > 1 ? item.SubItems[1].Text : string.Empty;
                        ShowEditVolunteerDialog(surname, email);
                    };
                    contextMenu.Items.Add(editItem);

                    var deleteItem = new ToolStripMenuItem("Elimina");
                    deleteItem.Click += (s, args) =>
                    {
                        var surname = item.Text;
                        _controller.OnDeleteVolunteer(surname);
                    };
                    contextMenu.Items.Add(deleteItem);

                    contextMenu.Show(lstVolunteers, e.Location);
                }
            }
        }

        /// <summary>
        /// Shows a dialog (styled like "Aggiungi Contatto") pre-filled with the
        /// selected volunteer's current surname/email, and saves the changes via
        /// <see cref="IVolunteerNotificationController.OnEditVolunteer"/>.
        /// </summary>
        /// <param name="currentSurname">The volunteer's current surname</param>
        /// <param name="currentEmail">The volunteer's current email</param>
        private void ShowEditVolunteerDialog(string currentSurname, string currentEmail)
        {
            using (var form = new Form())
            {
                form.Text = "Modifica Contatto";
                form.Size = new Size(420, 210);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.BackColor = ThemeManager.ColorAppBackground;

                var lblSurname = new Label { Text = "Cognome:", Location = new Point(24, 28), Size = new Size(100, 20), ForeColor = ThemeManager.ColorPrimary, Font = new Font("Segoe UI", 9F) };
                var txtSurname = new Controls.ModernTextBox { Location = new Point(134, 24), Size = new Size(250, 28), ForeColor = Color.Black, Text = currentSurname };

                var lblEmail = new Label { Text = "Email:", Location = new Point(24, 70), Size = new Size(100, 20), ForeColor = ThemeManager.ColorPrimary, Font = new Font("Segoe UI", 9F) };
                var txtEmail = new Controls.ModernTextBox { Location = new Point(134, 66), Size = new Size(250, 28), ForeColor = Color.Black, Text = currentEmail };

                var btnOk = new Controls.ModernButton { Text = "Salva", Location = new Point(134, 118), Size = new Size(115, 38), DialogResult = DialogResult.OK };
                ThemeManager.ApplyPrimary(btnOk);
                var btnCancel = new Controls.ModernButton { Text = "Annulla", Location = new Point(269, 118), Size = new Size(115, 38), DialogResult = DialogResult.Cancel };
                ThemeManager.ApplyAccent(btnCancel);

                form.Controls.AddRange(new Control[] { lblSurname, txtSurname, lblEmail, txtEmail, btnOk, btnCancel });
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;

                if (form.ShowDialog() == DialogResult.OK)
                {
                    _controller.OnEditVolunteer(currentSurname, txtSurname.Text, txtEmail.Text);
                }
            }
        }
        
        /// <summary>
        /// Handles the Add Contact button click event.
        /// Prompts user to enter surname and email for new volunteer.
        /// </summary>
        private void BtnAddContact_Click(object? sender, EventArgs e)
        {
            using (var form = new Form())
            {
                form.Text = "Aggiungi Contatto";
                form.Size = new Size(420, 210);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                form.BackColor = ThemeManager.ColorAppBackground;

                var lblSurname = new Label { Text = "Cognome:", Location = new Point(24, 28), Size = new Size(100, 20), ForeColor = ThemeManager.ColorPrimary, Font = new Font("Segoe UI", 9F) };
                var txtSurname = new Controls.ModernTextBox { Location = new Point(134, 24), Size = new Size(250, 28), ForeColor = Color.Black };

                var lblEmail = new Label { Text = "Email:", Location = new Point(24, 70), Size = new Size(100, 20), ForeColor = ThemeManager.ColorPrimary, Font = new Font("Segoe UI", 9F) };
                var txtEmail = new Controls.ModernTextBox { Location = new Point(134, 66), Size = new Size(250, 28), ForeColor = Color.Black };

                var btnOk = new Controls.ModernButton { Text = "OK", Location = new Point(134, 118), Size = new Size(115, 38), DialogResult = DialogResult.OK };
                ThemeManager.ApplyPrimary(btnOk);
                var btnCancel = new Controls.ModernButton { Text = "Annulla", Location = new Point(269, 118), Size = new Size(115, 38), DialogResult = DialogResult.Cancel };
                ThemeManager.ApplyAccent(btnCancel);

                form.Controls.AddRange(new Control[] { lblSurname, txtSurname, lblEmail, txtEmail, btnOk, btnCancel });
                form.AcceptButton = btnOk;
                form.CancelButton = btnCancel;
                
                if (form.ShowDialog() == DialogResult.OK)
                {
                    _controller.OnAddVolunteer(txtSurname.Text, txtEmail.Text);
                }
            }
        }
        
        /// <summary>
        /// Handles the Delete All button click event.
        /// </summary>
        private void BtnDeleteAll_Click(object? sender, EventArgs e)
        {
            _controller.OnDeleteAllVolunteers();
        }
        
        /// <summary>
        /// Handles the Clear Credentials button click event.
        /// </summary>
        private void BtnClearCredentials_Click(object? sender, EventArgs e)
        {
            _controller.OnClearGmailCredentials();
        }
        
        /// <summary>
        /// Handles Gmail email textbox text changed event.
        /// </summary>

        private void BtnSaveCredentials_Click(object? sender, EventArgs e)
        {
            _controller.OnGmailCredentialsUpdated(txtGmailEmail.Text, txtGmailPassword.Text);
            _controller.SaveGmailCredentials();
            MessageBox.Show("Credenziali salvate.", "Salva Credenziali", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void BtnImportCredentials_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Importa Credenziali Gmail";
                dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var json = System.IO.File.ReadAllText(dialog.FileName);
                        using var doc = System.Text.Json.JsonDocument.Parse(json);
                        var root = doc.RootElement;
                        var email = root.GetProperty("email").GetString() ?? string.Empty;
                        var password = root.GetProperty("password").GetString() ?? string.Empty;
                        _controller.OnGmailCredentialsUpdated(email, password);
                        DisplayGmailCredentials(email, password);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Errore nell'importazione: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        /// <summary>
        /// Handles the Export Credentials button click event.
        /// Writes the current Gmail credentials to a JSON file in the same
        /// {"email": ..., "password": ...} format read by "Importa", so files
        /// exported here can be re-imported later (or on another machine).
        /// </summary>
        private void BtnExportCredentials_Click(object? sender, EventArgs e)
        {
            using (var dialog = new SaveFileDialog())
            {
                dialog.Title = "Esporta Credenziali Gmail";
                dialog.Filter = "JSON Files (*.json)|*.json|All Files (*.*)|*.*";
                dialog.FileName = "credenziali-gmail.json";
                dialog.RestoreDirectory = true;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        var (email, appPassword) = _controller.GetGmailCredentials();
                        var payload = new
                        {
                            email,
                            password = appPassword
                        };
                        var json = System.Text.Json.JsonSerializer.Serialize(payload, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
                        System.IO.File.WriteAllText(dialog.FileName, json);
                        MessageBox.Show("Credenziali esportate.", "Esporta Credenziali", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Errore nell'esportazione: " + ex.Message, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }
        private void TxtGmailEmail_TextChanged(object? sender, EventArgs e)
        {
            _controller.OnGmailCredentialsUpdated(txtGmailEmail.Text, txtGmailPassword.Text);
        }
        
        /// <summary>
        /// Handles Gmail password textbox text changed event.
        /// </summary>
        private void TxtGmailPassword_TextChanged(object? sender, EventArgs e)
        {
            _controller.OnGmailCredentialsUpdated(txtGmailEmail.Text, txtGmailPassword.Text);
        }
        
        /// <summary>
        /// Handles the Select Excel button click event.
        /// </summary>
        private void BtnSelectExcel_Click(object? sender, EventArgs e)
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Title = "Seleziona file Excel";
                dialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*";
                dialog.FilterIndex = 1;
                dialog.RestoreDirectory = true;
                
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    _controller.OnNotificationExcelFileSelected(dialog.FileName);
                }
            }
        }
        
        /// <summary>
        /// Handles the sheet combobox selection changed event.
        /// </summary>
        private void CmbSheets_SelectedIndexChanged(object? sender, EventArgs e)
        {
            if (cmbSheets.SelectedItem != null)
            {
                _controller.OnSheetSelected(cmbSheets.SelectedItem.ToString()!);
            }
        }
        
        /// <summary>
        /// Handles the Send Emails button click event.
        /// </summary>
        private async void BtnSendEmails_Click(object? sender, EventArgs e)
        {
            await _controller.OnSendEmailsClickedAsync();
        }
        
        #endregion
        
        #region IVolunteerUI Implementation
        
        /// <summary>
        /// Displays the list of volunteer contacts.
        /// Validates: Requirements 8.1, 8.9
        /// </summary>
        /// <param name="volunteers">Dictionary of surname to email mappings</param>
        public void DisplayVolunteerList(Dictionary<string, string> volunteers)
        {
            if (lstVolunteers.InvokeRequired)
            {
                lstVolunteers.Invoke(new Action(() => DisplayVolunteerList(volunteers)));
                return;
            }
            
            lstVolunteers.Items.Clear();
            
            foreach (var volunteer in volunteers.OrderBy(v => v.Key))
            {
                var item = new ListViewItem(volunteer.Key);
                item.SubItems.Add(volunteer.Value);
                lstVolunteers.Items.Add(item);
            }

            // A successful refresh means the last volunteer-list operation (add,
            // delete, import) succeeded, so clear any previously shown error.
            lblVolunteerError.Text = string.Empty;
        }
        
        /// <summary>
        /// Displays Gmail credentials in the UI.
        /// Validates: Requirements 3.3
        /// </summary>
        /// <param name="email">Gmail email address</param>
        /// <param name="password">Gmail app password</param>
        public void DisplayGmailCredentials(string email, string password)
        {
            if (txtGmailEmail.InvokeRequired)
            {
                txtGmailEmail.Invoke(new Action(() => DisplayGmailCredentials(email, password)));
                return;
            }
            
            // Temporarily disable event handlers to avoid triggering save during load
            txtGmailEmail.TextChanged -= TxtGmailEmail_TextChanged;
            txtGmailPassword.TextChanged -= TxtGmailPassword_TextChanged;
            
            txtGmailEmail.Text = email ?? string.Empty;
            txtGmailPassword.Text = password ?? string.Empty;
            
            // Re-enable event handlers
            txtGmailEmail.TextChanged += TxtGmailEmail_TextChanged;
            txtGmailPassword.TextChanged += TxtGmailPassword_TextChanged;
        }
        
        /// <summary>
        /// Displays available sheet names from Excel file.
        /// Validates: Requirements 2.3
        /// </summary>
        /// <param name="sheetNames">List of sheet names</param>
        public void DisplaySheetNames(List<string> sheetNames)
        {
            if (cmbSheets.InvokeRequired)
            {
                cmbSheets.Invoke(new Action(() => DisplaySheetNames(sheetNames)));
                return;
            }
            
            cmbSheets.Items.Clear();
            foreach (var sheetName in sheetNames)
            {
                cmbSheets.Items.Add(sheetName);
            }
            
            if (cmbSheets.Items.Count > 0)
            {
                cmbSheets.SelectedIndex = 0;
            }
        }
        
        /// <summary>
        /// Enables or disables the send emails button.
        /// Validates: Requirements 5.1
        /// </summary>
        /// <param name="enabled">True to enable, false to disable</param>
        public void EnableSendEmailsButton(bool enabled)
        {
            if (btnSendEmails.InvokeRequired)
            {
                btnSendEmails.Invoke(new Action(() => btnSendEmails.Enabled = enabled));
            }
            else
            {
                btnSendEmails.Enabled = enabled;
            }
        }
        
        /// <summary>
        /// Shows email sending progress.
        /// Validates: Requirements 5.7
        /// </summary>
        /// <param name="message">Progress message</param>
        public void ShowEmailProgress(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = ThemeManager.ColorSecondary;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = ThemeManager.ColorSecondary;
            }
        }
        
        /// <summary>
        /// Shows email sending summary.
        /// Validates: Requirements 5.7
        /// </summary>
        /// <param name="successCount">Number of successful sends</param>
        /// <param name="failureCount">Number of failed sends</param>
        public void ShowEmailSummary(int successCount, int failureCount)
        {
            var message = string.Format(Properties.Resources.EmailSummaryTemplate, successCount, failureCount);
            
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = failureCount > 0 ? Color.FromArgb(0xE0, 0x8E, 0x00) : ThemeManager.ColorPrimary;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = failureCount > 0 ? Color.FromArgb(0xE0, 0x8E, 0x00) : ThemeManager.ColorPrimary;
            }
        }
        
        /// <summary>
        /// Prompts user for confirmation.
        /// Validates: Requirements 8.7
        /// </summary>
        /// <param name="message">Confirmation message</param>
        /// <returns>True if confirmed, false otherwise</returns>
        public bool ConfirmAction(string message)
        {
            var result = MessageBox.Show(message, "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            return result == DialogResult.Yes;
        }
        
        /// <summary>
        /// Shows an error message to the user.
        /// </summary>
        /// <param name="message">Error message to display</param>
        public void ShowErrorMessage(string message)
        {
            if (lblStatus.InvokeRequired)
            {
                lblStatus.Invoke(new Action(() =>
                {
                    lblStatus.Text = message;
                    lblStatus.ForeColor = ThemeManager.ColorError;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = ThemeManager.ColorError;
            }
        }

        /// <summary>
        /// Shows an error message scoped to the "Elenco Volontari" section (invalid
        /// surname/email, import/delete failures), displayed under the section's
        /// subtitle rather than the general "Invio Email" status label.
        /// </summary>
        /// <param name="message">Error message to display</param>
        public void ShowVolunteerErrorMessage(string message)
        {
            if (lblVolunteerError.InvokeRequired)
            {
                lblVolunteerError.Invoke(new Action(() =>
                {
                    lblVolunteerError.Text = message;
                }));
            }
            else
            {
                lblVolunteerError.Text = message;
            }
        }
        
        #endregion
    }
}
