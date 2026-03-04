using NUnit.Framework;
using Moq;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests;

/// <summary>
/// Unit tests for VolunteerPanel UI component.
/// Validates: Requirements 1.1, 1.2, 8.1, 8.2, 8.3, 8.6, 8.9
/// </summary>
[TestFixture]
[Apartment(System.Threading.ApartmentState.STA)]
public class VolunteerPanelTests
{
    private Mock<IVolunteerNotificationController> _mockController = null!;
    private VolunteerPanel _panel = null!;
    
    [SetUp]
    public void Setup()
    {
        _mockController = new Mock<IVolunteerNotificationController>();
        _panel = new VolunteerPanel(_mockController.Object);
    }
    
    [TearDown]
    public void TearDown()
    {
        _panel?.Dispose();
    }
    
    #region Button Label Tests (Requirements 1.1, 8.2, 8.6)
    
    [Test]
    public void AddVolunteersButton_ShouldHaveCorrectLabel()
    {
        // Act
        var button = FindControl<Button>(_panel, "btnAddVolunteers");
        
        // Assert
        Assert.That(button, Is.Not.Null, "Add Volunteers button should exist");
        Assert.That(button!.Text, Is.EqualTo(Properties.Resources.AddVolunteersButton));
    }
    
    [Test]
    public void AddContactButton_ShouldHaveCorrectLabel()
    {
        // Act
        var button = FindControl<Button>(_panel, "btnAddContact");
        
        // Assert
        Assert.That(button, Is.Not.Null, "Add Contact button should exist");
        Assert.That(button!.Text, Is.EqualTo(Properties.Resources.AddContactButton));
    }
    
    [Test]
    public void DeleteAllButton_ShouldHaveCorrectLabel()
    {
        // Act
        var button = FindControl<Button>(_panel, "btnDeleteAll");
        
        // Assert
        Assert.That(button, Is.Not.Null, "Delete All button should exist");
        Assert.That(button!.Text, Is.EqualTo(Properties.Resources.DeleteAllButton));
    }
    
    [Test]
    public void SendEmailsButton_ShouldHaveCorrectLabel()
    {
        // Act
        var button = FindControl<Button>(_panel, "btnSendEmails");
        
        // Assert
        Assert.That(button, Is.Not.Null, "Send Emails button should exist");
        Assert.That(button!.Text, Is.EqualTo(Properties.Resources.SendEmailsButton));
    }
    
    #endregion
    
    #region Volunteer List Display Tests (Property 18, Requirement 8.1)
    
    [Test]
    public void DisplayVolunteerList_WithEmptyDictionary_ShouldShowNoItems()
    {
        // Arrange
        var volunteers = new Dictionary<string, string>();
        
        // Act
        _panel.DisplayVolunteerList(volunteers);
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        Assert.That(listView!.Items.Count, Is.EqualTo(0));
    }
    
    [Test]
    public void DisplayVolunteerList_WithVolunteers_ShouldShowAllEntries()
    {
        // Arrange
        var volunteers = new Dictionary<string, string>
        {
            { "Rossi", "rossi@example.com" },
            { "Bianchi", "bianchi@example.com" },
            { "Verdi", "verdi@example.com" }
        };
        
        // Act
        _panel.DisplayVolunteerList(volunteers);
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        Assert.That(listView!.Items.Count, Is.EqualTo(3));
    }
    
    [Test]
    public void DisplayVolunteerList_ShouldShowSurnameAndEmail()
    {
        // Arrange
        var volunteers = new Dictionary<string, string>
        {
            { "Rossi", "rossi@example.com" }
        };
        
        // Act
        _panel.DisplayVolunteerList(volunteers);
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        Assert.That(listView!.Items.Count, Is.EqualTo(1));
        Assert.That(listView.Items[0].Text, Is.EqualTo("Rossi"));
        Assert.That(listView.Items[0].SubItems[1].Text, Is.EqualTo("rossi@example.com"));
    }
    
    [Test]
    public void DisplayVolunteerList_ShouldSortBySurname()
    {
        // Arrange
        var volunteers = new Dictionary<string, string>
        {
            { "Verdi", "verdi@example.com" },
            { "Bianchi", "bianchi@example.com" },
            { "Rossi", "rossi@example.com" }
        };
        
        // Act
        _panel.DisplayVolunteerList(volunteers);
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        Assert.That(listView!.Items.Count, Is.EqualTo(3));
        Assert.That(listView.Items[0].Text, Is.EqualTo("Bianchi"));
        Assert.That(listView.Items[1].Text, Is.EqualTo("Rossi"));
        Assert.That(listView.Items[2].Text, Is.EqualTo("Verdi"));
    }
    
    #endregion
    
    #region Delete Button Presence Tests (Property 20, Requirement 8.9)
    
    [Test]
    public void DisplayVolunteerList_EachVolunteer_ShouldHaveDeleteButton()
    {
        // Arrange
        var volunteers = new Dictionary<string, string>
        {
            { "Rossi", "rossi@example.com" },
            { "Bianchi", "bianchi@example.com" }
        };
        
        // Act
        _panel.DisplayVolunteerList(volunteers);
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        foreach (ListViewItem item in listView!.Items)
        {
            Assert.That(item.Tag, Is.Not.Null, "Each item should have a delete button");
            Assert.That(item.Tag, Is.InstanceOf<Button>());
            var button = (Button)item.Tag;
            Assert.That(button.Text, Is.EqualTo("Elimina"));
        }
    }
    
    #endregion
    
    #region Sheet Display Tests (Requirement 2.3)
    
    [Test]
    public void DisplaySheetNames_WithSheets_ShouldPopulateComboBox()
    {
        // Arrange
        var sheetNames = new List<string> { "Sheet1", "Sheet2", "Sheet3" };
        
        // Act
        _panel.DisplaySheetNames(sheetNames);
        var comboBox = FindControl<ComboBox>(_panel, "cmbSheets");
        
        // Assert
        Assert.That(comboBox, Is.Not.Null);
        Assert.That(comboBox!.Items.Count, Is.EqualTo(3));
        Assert.That(comboBox.Items[0], Is.EqualTo("Sheet1"));
        Assert.That(comboBox.Items[1], Is.EqualTo("Sheet2"));
        Assert.That(comboBox.Items[2], Is.EqualTo("Sheet3"));
    }
    
    [Test]
    public void DisplaySheetNames_WithSheets_ShouldSelectFirstSheet()
    {
        // Arrange
        var sheetNames = new List<string> { "Sheet1", "Sheet2" };
        
        // Act
        _panel.DisplaySheetNames(sheetNames);
        var comboBox = FindControl<ComboBox>(_panel, "cmbSheets");
        
        // Assert
        Assert.That(comboBox, Is.Not.Null);
        Assert.That(comboBox!.SelectedIndex, Is.EqualTo(0));
        Assert.That(comboBox.SelectedItem, Is.EqualTo("Sheet1"));
    }
    
    [Test]
    public void DisplaySheetNames_WithEmptyList_ShouldClearComboBox()
    {
        // Arrange
        var sheetNames = new List<string>();
        
        // Act
        _panel.DisplaySheetNames(sheetNames);
        var comboBox = FindControl<ComboBox>(_panel, "cmbSheets");
        
        // Assert
        Assert.That(comboBox, Is.Not.Null);
        Assert.That(comboBox!.Items.Count, Is.EqualTo(0));
    }
    
    #endregion
    
    #region Button Enable/Disable Tests (Requirement 5.1)
    
    [Test]
    public void EnableSendEmailsButton_WithTrue_ShouldEnableButton()
    {
        // Act
        _panel.EnableSendEmailsButton(true);
        var button = FindControl<Button>(_panel, "btnSendEmails");
        
        // Assert
        Assert.That(button, Is.Not.Null);
        Assert.That(button!.Enabled, Is.True);
    }
    
    [Test]
    public void EnableSendEmailsButton_WithFalse_ShouldDisableButton()
    {
        // Act
        _panel.EnableSendEmailsButton(false);
        var button = FindControl<Button>(_panel, "btnSendEmails");
        
        // Assert
        Assert.That(button, Is.Not.Null);
        Assert.That(button!.Enabled, Is.False);
    }
    
    [Test]
    public void SendEmailsButton_InitialState_ShouldBeDisabled()
    {
        // Act
        var button = FindControl<Button>(_panel, "btnSendEmails");
        
        // Assert
        Assert.That(button, Is.Not.Null);
        Assert.That(button!.Enabled, Is.False);
    }
    
    #endregion
    
    #region Progress and Status Tests (Requirement 5.7)
    
    [Test]
    public void ShowEmailProgress_ShouldDisplayMessage()
    {
        // Arrange
        var message = "Invio email in corso...";
        
        // Act
        _panel.ShowEmailProgress(message);
        var label = FindControl<Label>(_panel, "lblStatus");
        
        // Assert
        Assert.That(label, Is.Not.Null);
        Assert.That(label!.Text, Is.EqualTo(message));
    }
    
    [Test]
    public void ShowEmailSummary_ShouldDisplayFormattedMessage()
    {
        // Arrange
        int successCount = 5;
        int failureCount = 2;
        
        // Act
        _panel.ShowEmailSummary(successCount, failureCount);
        var label = FindControl<Label>(_panel, "lblStatus");
        
        // Assert
        Assert.That(label, Is.Not.Null);
        Assert.That(label!.Text, Does.Contain("5"));
        Assert.That(label.Text, Does.Contain("2"));
    }
    
    [Test]
    public void ShowErrorMessage_ShouldDisplayMessage()
    {
        // Arrange
        var message = "Errore di test";
        
        // Act
        _panel.ShowErrorMessage(message);
        var label = FindControl<Label>(_panel, "lblStatus");
        
        // Assert
        Assert.That(label, Is.Not.Null);
        Assert.That(label!.Text, Is.EqualTo(message));
    }
    
    #endregion
    
    #region UI Component Existence Tests
    
    [Test]
    public void VolunteerPanel_ShouldHaveVolunteerContactsGroupBox()
    {
        // Act
        var groupBox = FindControl<GroupBox>(_panel, "grpVolunteerContacts");
        
        // Assert
        Assert.That(groupBox, Is.Not.Null);
        Assert.That(groupBox!.Text, Is.EqualTo(Properties.Resources.VolunteerContactsGroupTitle));
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveGmailCredentialsGroupBox()
    {
        // Act
        var groupBox = FindControl<GroupBox>(_panel, "grpGmailCredentials");
        
        // Assert
        Assert.That(groupBox, Is.Not.Null);
        Assert.That(groupBox!.Text, Is.EqualTo("Credenziali Gmail"));
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveExcelSelectionGroupBox()
    {
        // Act
        var groupBox = FindControl<GroupBox>(_panel, "grpExcelSelection");
        
        // Assert
        Assert.That(groupBox, Is.Not.Null);
        Assert.That(groupBox!.Text, Is.EqualTo("Selezione File Excel"));
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveEmailSendingGroupBox()
    {
        // Act
        var groupBox = FindControl<GroupBox>(_panel, "grpEmailSending");
        
        // Assert
        Assert.That(groupBox, Is.Not.Null);
        Assert.That(groupBox!.Text, Is.EqualTo("Invio Email"));
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveVolunteerListView()
    {
        // Act
        var listView = FindControl<ListView>(_panel, "lstVolunteers");
        
        // Assert
        Assert.That(listView, Is.Not.Null);
        Assert.That(listView!.Columns.Count, Is.GreaterThanOrEqualTo(2));
        Assert.That(listView.Columns[0].Text, Is.EqualTo(Properties.Resources.VolunteerListColumnSurname));
        Assert.That(listView.Columns[1].Text, Is.EqualTo(Properties.Resources.VolunteerListColumnEmail));
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveGmailEmailTextBox()
    {
        // Act
        var textBox = FindControl<TextBox>(_panel, "txtGmailEmail");
        
        // Assert
        Assert.That(textBox, Is.Not.Null);
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveGmailPasswordTextBox()
    {
        // Act
        var textBox = FindControl<TextBox>(_panel, "txtGmailPassword");
        
        // Assert
        Assert.That(textBox, Is.Not.Null);
        Assert.That(textBox!.UseSystemPasswordChar, Is.True, "Password should be masked");
    }
    
    [Test]
    public void VolunteerPanel_ShouldHaveProgressBar()
    {
        // Act
        var progressBar = FindControl<ProgressBar>(_panel, "progressBar");
        
        // Assert
        Assert.That(progressBar, Is.Not.Null);
    }
    
    #endregion
    
    #region Helper Methods
    
    /// <summary>
    /// Recursively finds a control by field name in the control hierarchy.
    /// </summary>
    private T? FindControl<T>(Control parent, string fieldName) where T : Control
    {
        // Check if the parent itself matches
        var field = parent.GetType().GetField(fieldName, 
            System.Reflection.BindingFlags.NonPublic | 
            System.Reflection.BindingFlags.Instance);
        
        if (field != null && field.GetValue(parent) is T control)
        {
            return control;
        }
        
        // Recursively search children
        foreach (Control child in parent.Controls)
        {
            var found = FindControl<T>(child, fieldName);
            if (found != null)
                return found;
        }
        
        return null;
    }
    
    #endregion
}
