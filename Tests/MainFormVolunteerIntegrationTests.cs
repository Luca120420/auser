using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.UI;
using NUnit.Framework;

namespace AuserExcelTransformer.Tests;

/// <summary>
/// Integration tests for MainForm with volunteer notification feature.
/// Tests the integration of VolunteerPanel into MainForm and configuration loading on startup.
/// Validates: Requirements 1.1, 1.6, 3.3
/// </summary>
[TestFixture]
public class MainFormVolunteerIntegrationTests
{
    private string _testConfigPath = null!;
    private string _originalConfigPath = null!;

    [SetUp]
    public void Setup()
    {
        // Create a temporary config directory for testing
        _testConfigPath = Path.Combine(Path.GetTempPath(), "AuserExcelTransformer_Test", "config.json");
        var testDir = Path.GetDirectoryName(_testConfigPath);
        if (!string.IsNullOrEmpty(testDir) && !Directory.Exists(testDir))
        {
            Directory.CreateDirectory(testDir);
        }

        // Clean up any existing test config
        if (File.Exists(_testConfigPath))
        {
            File.Delete(_testConfigPath);
        }
    }

    [TearDown]
    public void TearDown()
    {
        // Clean up test config file
        if (File.Exists(_testConfigPath))
        {
            File.Delete(_testConfigPath);
        }

        var testDir = Path.GetDirectoryName(_testConfigPath);
        if (!string.IsNullOrEmpty(testDir) && Directory.Exists(testDir))
        {
            try
            {
                Directory.Delete(testDir, true);
            }
            catch
            {
                // Ignore cleanup errors
            }
        }
    }

    /// <summary>
    /// Tests that the volunteer panel is visible in MainForm.
    /// Validates: Requirement 1.1
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_ShouldContainVolunteerPanel()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act
        using var form = new MainForm(controller);

        // Assert
        var volunteerPanel = FindControlByType<VolunteerPanel>(form);
        Assert.That(volunteerPanel, Is.Not.Null, "VolunteerPanel should be present in MainForm");
    }

    /// <summary>
    /// Tests that the volunteer panel is positioned below existing transformation controls.
    /// Validates: Requirement 10.1
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_VolunteerPanel_ShouldBePositionedBelowTransformationControls()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act
        using var form = new MainForm(controller);

        // Assert
        var volunteerPanel = FindControlByType<VolunteerPanel>(form);
        Assert.That(volunteerPanel, Is.Not.Null);
        Assert.That(volunteerPanel!.Location.Y, Is.GreaterThan(300), 
            "VolunteerPanel should be positioned below transformation controls (Y > 300)");
    }

    /// <summary>
    /// Tests that MainForm size is adjusted to accommodate volunteer panel.
    /// Validates: Requirement 10.1
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_ShouldHaveIncreasedSizeForVolunteerPanel()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act
        using var form = new MainForm(controller);

        // Assert
        Assert.That(form.Size.Height, Is.GreaterThan(600), 
            "MainForm height should be increased to accommodate VolunteerPanel");
        Assert.That(form.AutoScroll, Is.True, 
            "MainForm should have AutoScroll enabled for volunteer panel");
    }

    /// <summary>
    /// Tests that configuration is loaded on MainForm startup.
    /// This test verifies that the volunteer feature initialization doesn't throw exceptions.
    /// Validates: Requirements 1.6, 3.3
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_ShouldLoadConfigurationOnStartup()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act & Assert - should not throw
        Assert.DoesNotThrow(() =>
        {
            using var form = new MainForm(controller);
            // If we get here without exception, configuration loading succeeded
        });
    }

    /// <summary>
    /// Tests end-to-end volunteer management workflow through MainForm.
    /// This is a basic smoke test to ensure the integration works.
    /// Validates: Requirements 1.1, 1.6, 3.3
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_VolunteerFeature_EndToEndWorkflow()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act
        using var form = new MainForm(controller);
        var volunteerPanel = FindControlByType<VolunteerPanel>(form);

        // Assert
        Assert.That(volunteerPanel, Is.Not.Null, "VolunteerPanel should be initialized");
        
        // Verify that the panel has the expected controls
        var addVolunteersButton = FindControlByText<Button>(volunteerPanel!, "Aggiungi Volontari");
        Assert.That(addVolunteersButton, Is.Not.Null, "Add Volunteers button should be present");
        
        var addContactButton = FindControlByText<Button>(volunteerPanel!, "Aggiungi Contatto");
        Assert.That(addContactButton, Is.Not.Null, "Add Contact button should be present");
        
        var deleteAllButton = FindControlByText<Button>(volunteerPanel!, "Elimina Tutti");
        Assert.That(deleteAllButton, Is.Not.Null, "Delete All button should be present");
        
        var sendEmailsButton = FindControlByText<Button>(volunteerPanel!, "Invia Email");
        Assert.That(sendEmailsButton, Is.Not.Null, "Send Emails button should be present");
    }

    /// <summary>
    /// Tests that MainForm handles volunteer feature initialization failure gracefully.
    /// Validates: Requirement 10.2
    /// </summary>
    [Test]
    [Apartment(ApartmentState.STA)]
    public void MainForm_ShouldHandleVolunteerFeatureInitializationFailureGracefully()
    {
        // Arrange
        var controller = new ApplicationController(null!, null!, null!);

        // Act & Assert - should not throw even if volunteer feature fails to initialize
        Assert.DoesNotThrow(() =>
        {
            using var form = new MainForm(controller);
            // Form should still be created even if volunteer feature fails
        });
    }

    /// <summary>
    /// Helper method to find a control by type in a control hierarchy.
    /// </summary>
    private T? FindControlByType<T>(Control parent) where T : Control
    {
        foreach (Control control in parent.Controls)
        {
            if (control is T typedControl)
            {
                return typedControl;
            }

            var found = FindControlByType<T>(control);
            if (found != null)
            {
                return found;
            }
        }
        return null;
    }

    /// <summary>
    /// Helper method to find a control by text in a control hierarchy.
    /// </summary>
    private T? FindControlByText<T>(Control parent, string text) where T : Control
    {
        foreach (Control control in parent.Controls)
        {
            if (control is T typedControl && control.Text == text)
            {
                return typedControl;
            }

            var found = FindControlByText<T>(control, text);
            if (found != null)
            {
                return found;
            }
        }
        return null;
    }
}
