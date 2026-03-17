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
        
        // Volunteer Contacts Section
        private GroupBox grpVolunteerContacts = null!;
        private ListView lstVolunteers = null!;
        private Button btnAddVolunteers = null!;
        private Button btnAddContact = null!;
        private Button btnDeleteAll = null!;
        
        // Gmail Credentials Section
        private GroupBox grpGmailCredentials = null!;
        private Label lblGmailEmail = null!;
        private TextBox txtGmailEmail = null!;
        private Label lblGmailPassword = null!;
        private TextBox txtGmailPassword = null!;
        private Button btnClearCredentials = null!;
        private Button btnSaveCredentials = null!;
        private Button btnImportCredentials = null!;
        
        // Excel Selection Section
        private GroupBox grpExcelSelection = null!;
        private Button btnSelectExcel = null!;
        private Label lblSheet = null!;
        private ComboBox cmbSheets = null!;
        
        // Email Sending Section
        private GroupBox grpEmailSending = null!;
        private Button btnSendEmails = null!;
        private ProgressBar progressBar = null!;
        private Label lblStatus = null!;
        
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
            this.Size = new Size(800, 600);
            this.AutoScroll = false; // Disable panel scrolling - let MainForm handle scrolling
            
            InitializeVolunteerContactsSection();
            InitializeGmailCredentialsSection();
            InitializeExcelSelectionSection();
            InitializeEmailSendingSection();
        }
        
        /// <summary>
        /// Initializes the volunteer contacts section.
        /// Validates: Requirements 1.1, 8.1, 8.2, 8.6, 8.9
        /// </summary>
        private void InitializeVolunteerContactsSection()
        {
            grpVolunteerContacts = new GroupBox
            {
                Text = Properties.Resources.VolunteerContactsGroupTitle,
                Location = new Point(10, 10),
                Size = new Size(780, 200),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            
            // ListView for volunteer contacts
            lstVolunteers = new ListView
            {
                Location = new Point(10, 25),
                Size = new Size(760, 120),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                Font = new Font("Segoe UI", 9F)
            };
            
            // Add columns
            lstVolunteers.Columns.Add(Properties.Resources.VolunteerListColumnSurname, 300);
            lstVolunteers.Columns.Add(Properties.Resources.VolunteerListColumnEmail, 400);
            
            // Handle mouse click for delete action
            lstVolunteers.MouseClick += LstVolunteers_MouseClick;
            
            grpVolunteerContacts.Controls.Add(lstVolunteers);
            
            // Add Volunteers button
            btnAddVolunteers = new Button
            {
                Text = Properties.Resources.AddVolunteersButton,
                Location = new Point(10, 155),
                Size = new Size(180, 35),
                Font = new Font("Segoe UI", 9F)
            };
            btnAddVolunteers.Click += BtnAddVolunteers_Click;
            grpVolunteerContacts.Controls.Add(btnAddVolunteers);
            
            // Add Contact button
            btnAddContact = new Button
            {
                Text = Properties.Resources.AddContactButton,
                Location = new Point(200, 155),
                Size = new Size(180, 35),
                Font = new Font("Segoe UI", 9F)
            };
            btnAddContact.Click += BtnAddContact_Click;
            grpVolunteerContacts.Controls.Add(btnAddContact);
            
            // Delete All button
            btnDeleteAll = new Button
            {
                Text = Properties.Resources.DeleteAllButton,
                Location = new Point(390, 155),
                Size = new Size(180, 35),
                Font = new Font("Segoe UI", 9F)
            };
            btnDeleteAll.Click += BtnDeleteAll_Click;
            grpVolunteerContacts.Controls.Add(btnDeleteAll);
            
            this.Controls.Add(grpVolunteerContacts);
        }
        
        /// <summary>
        /// Initializes the Gmail credentials section.
        /// Validates: Requirements 3.1
        /// </summary>
        private void InitializeGmailCredentialsSection()
        {
            grpGmailCredentials = new GroupBox
            {
                Text = "Credenziali Gmail",
                Location = new Point(10, 220),
                Size = new Size(780, 140),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            
            // Gmail email label
            lblGmailEmail = new Label
            {
                Text = "Email Gmail:",
                Location = new Point(10, 30),
                Size = new Size(120, 20),
                Font = new Font("Segoe UI", 9F)
            };
            grpGmailCredentials.Controls.Add(lblGmailEmail);
            
            // Gmail email textbox
            txtGmailEmail = new TextBox
            {
                Location = new Point(140, 28),
                Size = new Size(620, 25),
                Font = new Font("Segoe UI", 9F)
            };
            txtGmailEmail.TextChanged += TxtGmailEmail_TextChanged;
            grpGmailCredentials.Controls.Add(txtGmailEmail);
            
            // Gmail password label
            lblGmailPassword = new Label
            {
                Text = "Password App:",
                Location = new Point(10, 65),
                Size = new Size(120, 20),
                Font = new Font("Segoe UI", 9F)
            };
            grpGmailCredentials.Controls.Add(lblGmailPassword);
            
            // Gmail password textbox
            txtGmailPassword = new TextBox
            {
                Location = new Point(140, 63),
                Size = new Size(620, 25),
                Font = new Font("Segoe UI", 9F),
                UseSystemPasswordChar = true
            };
            txtGmailPassword.TextChanged += TxtGmailPassword_TextChanged;
            grpGmailCredentials.Controls.Add(txtGmailPassword);
            
            // Clear Credentials button
            btnClearCredentials = new Button
            {
                Text = "Cancella Credenziali",
                Location = new Point(10, 95),
                Size = new Size(180, 30),
                Font = new Font("Segoe UI", 9F)
            };
            btnClearCredentials.Click += BtnClearCredentials_Click;
            grpGmailCredentials.Controls.Add(btnClearCredentials);
            btnSaveCredentials = new Button
            {
                Text = "Salva Credenziali",
                Location = new Point(210, 95),
                Size = new Size(150, 30),
                Font = new Font("Segoe UI", 9F)
            };
            btnSaveCredentials.Click += BtnSaveCredentials_Click;
            grpGmailCredentials.Controls.Add(btnSaveCredentials);

            btnImportCredentials = new Button
            {
                Text = "Importa Credenziali",
                Location = new Point(385, 95),
                Size = new Size(160, 30),
                Font = new Font("Segoe UI", 9F)
            };
            btnImportCredentials.Click += BtnImportCredentials_Click;
            grpGmailCredentials.Controls.Add(btnImportCredentials);
            
            this.Controls.Add(grpGmailCredentials);
        }
        
        /// <summary>
        /// Initializes the Excel selection section.
        /// Validates: Requirements 2.1, 2.3, 2.4
        /// </summary>
        private void InitializeExcelSelectionSection()
        {
            grpExcelSelection = new GroupBox
            {
                Text = "Selezione File Excel",
                Location = new Point(10, 360),
                Size = new Size(780, 100),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            
            // Select Excel button
            btnSelectExcel = new Button
            {
                Text = "Seleziona File Excel",
                Location = new Point(10, 30),
                Size = new Size(200, 35),
                Font = new Font("Segoe UI", 9F)
            };
            btnSelectExcel.Click += BtnSelectExcel_Click;
            grpExcelSelection.Controls.Add(btnSelectExcel);
            
            // Sheet label
            lblSheet = new Label
            {
                Text = "Foglio:",
                Location = new Point(10, 70),
                Size = new Size(80, 20),
                Font = new Font("Segoe UI", 9F)
            };
            grpExcelSelection.Controls.Add(lblSheet);
            
            // Sheet combobox
            cmbSheets = new ComboBox
            {
                Location = new Point(100, 68),
                Size = new Size(660, 25),
                Font = new Font("Segoe UI", 9F),
                DropDownStyle = ComboBoxStyle.DropDownList
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
            grpEmailSending = new GroupBox
            {
                Text = "Invio Email",
                Location = new Point(10, 470),
                Size = new Size(780, 140),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold)
            };
            
            // Send Emails button
            btnSendEmails = new Button
            {
                Text = Properties.Resources.SendEmailsButton,
                Location = new Point(10, 30),
                Size = new Size(200, 40),
                Font = new Font("Segoe UI", 10F, FontStyle.Bold),
                Enabled = false
            };
            btnSendEmails.Click += BtnSendEmails_Click;
            grpEmailSending.Controls.Add(btnSendEmails);
            
            // Progress bar
            progressBar = new ProgressBar
            {
                Location = new Point(10, 80),
                Size = new Size(760, 20),
                Style = ProgressBarStyle.Continuous
            };
            grpEmailSending.Controls.Add(progressBar);
            
            // Status label
            lblStatus = new Label
            {
                Text = "",
                Location = new Point(10, 110),
                Size = new Size(760, 20),
                Font = new Font("Segoe UI", 9F)
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
        /// Handles mouse click on volunteer list for context menu.
        /// </summary>
        private void LstVolunteers_MouseClick(object? sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var item = lstVolunteers.GetItemAt(e.X, e.Y);
                if (item != null)
                {
                    // Show context menu for delete
                    var contextMenu = new ContextMenuStrip();
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
        /// Handles the Add Contact button click event.
        /// Prompts user to enter surname and email for new volunteer.
        /// </summary>
        private void BtnAddContact_Click(object? sender, EventArgs e)
        {
            using (var form = new Form())
            {
                form.Text = "Aggiungi Contatto";
                form.Size = new Size(400, 180);
                form.FormBorderStyle = FormBorderStyle.FixedDialog;
                form.StartPosition = FormStartPosition.CenterParent;
                form.MaximizeBox = false;
                form.MinimizeBox = false;
                
                var lblSurname = new Label { Text = "Cognome:", Location = new Point(20, 20), Size = new Size(100, 20) };
                var txtSurname = new TextBox { Location = new Point(130, 18), Size = new Size(240, 25) };
                
                var lblEmail = new Label { Text = "Email:", Location = new Point(20, 60), Size = new Size(100, 20) };
                var txtEmail = new TextBox { Location = new Point(130, 58), Size = new Size(240, 25) };
                
                var btnOk = new Button { Text = "OK", Location = new Point(130, 100), Size = new Size(100, 30), DialogResult = DialogResult.OK };
                var btnCancel = new Button { Text = "Annulla", Location = new Point(240, 100), Size = new Size(100, 30), DialogResult = DialogResult.Cancel };
                
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
                    lblStatus.ForeColor = Color.Blue;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.Blue;
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
                    lblStatus.ForeColor = failureCount > 0 ? Color.Orange : Color.Green;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = failureCount > 0 ? Color.Orange : Color.Green;
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
                    lblStatus.ForeColor = Color.Red;
                }));
            }
            else
            {
                lblStatus.Text = message;
                lblStatus.ForeColor = Color.Red;
            }
        }
        
        #endregion
    }
}
