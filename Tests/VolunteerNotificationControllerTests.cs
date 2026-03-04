using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using NUnit.Framework;
using Moq;
using AuserExcelTransformer.Services;
using AuserExcelTransformer.Models;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for VolunteerNotificationController class.
    /// Tests specific scenarios, edge cases, and state management.
    /// </summary>
    [TestFixture]
    public class VolunteerNotificationControllerTests
    {
        private Mock<IVolunteerManager> _mockVolunteerManager = null!;
        private Mock<IEmailService> _mockEmailService = null!;
        private Mock<IConfigurationService> _mockConfigurationService = null!;
        private Mock<IExcelManager> _mockExcelManager = null!;
        private Mock<IVolunteerUI> _mockUI = null!;
        private VolunteerNotificationController _controller = null!;

        [SetUp]
        public void Setup()
        {
            _mockVolunteerManager = new Mock<IVolunteerManager>();
            _mockEmailService = new Mock<IEmailService>();
            _mockConfigurationService = new Mock<IConfigurationService>();
            _mockExcelManager = new Mock<IExcelManager>();
            _mockUI = new Mock<IVolunteerUI>();

            // Setup default configuration to return empty config
            _mockConfigurationService
                .Setup(x => x.LoadConfiguration())
                .Returns(new AppConfiguration());
        }

        /// <summary>
        /// Test that configuration is loaded on controller construction.
        /// **Validates: Requirements 1.6, 3.3**
        /// </summary>
        [Test]
        public void Constructor_ShouldLoadConfiguration()
        {
            // Arrange
            var config = new AppConfiguration
            {
                VolunteerFilePath = "volunteers.json",
                GmailCredentials = new GmailCredentials
                {
                    Email = "test@gmail.com",
                    AppPassword = "testpassword"
                },
                LastExcelFilePath = "data.xlsx",
                LastSheetName = "Sheet1"
            };

            _mockConfigurationService
                .Setup(x => x.LoadConfiguration())
                .Returns(config);

            var volunteers = new Dictionary<string, string>
            {
                { "Rossi", "rossi@example.com" },
                { "Bianchi", "bianchi@example.com" }
            };

            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers("volunteers.json"))
                .Returns(volunteers);

            // Act
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Assert
            _mockConfigurationService.Verify(x => x.LoadConfiguration(), Times.Once,
                "Controller should load configuration on construction (Requirements 1.6, 3.3)");

            _mockVolunteerManager.Verify(x => x.LoadVolunteers("volunteers.json"), Times.Once,
                "Controller should load volunteers from configured file path on construction");

            _mockUI.Verify(x => x.DisplayVolunteerList(volunteers), Times.Once,
                "Controller should display loaded volunteers in UI on construction");
        }

        /// <summary>
        /// Test that configuration loading handles missing volunteer file gracefully.
        /// **Validates: Requirements 1.6, 3.3**
        /// </summary>
        [Test]
        public void Constructor_WithMissingVolunteerFile_ShouldStartWithEmptyVolunteers()
        {
            // Arrange
            var config = new AppConfiguration
            {
                VolunteerFilePath = "missing.json"
            };

            _mockConfigurationService
                .Setup(x => x.LoadConfiguration())
                .Returns(config);

            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers("missing.json"))
                .Throws(new System.IO.FileNotFoundException());

            // Act
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Assert - Should not throw, should start with empty volunteers
            var volunteers = _controller.GetVolunteers();
            Assert.That(volunteers.Count, Is.EqualTo(0),
                "Controller should start with empty volunteers when file loading fails");
        }

        /// <summary>
        /// Test CanSendEmails returns false when Gmail credentials are not configured.
        /// **Validates: Requirements 2.1, 3.6, 5.1**
        /// </summary>
        [Test]
        public void CanSendEmails_WithoutGmailCredentials_ShouldReturnFalse()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Add volunteers
            var volunteers = new Dictionary<string, string> { { "Rossi", "rossi@example.com" } };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");

            // Select sheet
            _controller.OnSheetSelected("Sheet1");

            // Act - Gmail credentials NOT configured
            bool canSend = _controller.CanSendEmails();

            // Assert
            Assert.That(canSend, Is.False,
                "CanSendEmails should return false when Gmail credentials are not configured (Requirements 3.6, 5.1)");
        }

        /// <summary>
        /// Test CanSendEmails returns false when no volunteers are loaded.
        /// **Validates: Requirements 2.1, 3.6, 5.1**
        /// </summary>
        [Test]
        public void CanSendEmails_WithoutVolunteers_ShouldReturnFalse()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Configure Gmail credentials
            _controller.OnGmailCredentialsUpdated("test@gmail.com", "password");

            // Select sheet
            _controller.OnSheetSelected("Sheet1");

            // Act - No volunteers loaded
            bool canSend = _controller.CanSendEmails();

            // Assert
            Assert.That(canSend, Is.False,
                "CanSendEmails should return false when no volunteers are loaded (Requirements 2.1, 5.1)");
        }

        /// <summary>
        /// Test CanSendEmails returns false when no sheet is selected.
        /// **Validates: Requirements 2.1, 3.6, 5.1**
        /// </summary>
        [Test]
        public void CanSendEmails_WithoutSheetSelected_ShouldReturnFalse()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Configure Gmail credentials
            _controller.OnGmailCredentialsUpdated("test@gmail.com", "password");

            // Add volunteers
            var volunteers = new Dictionary<string, string> { { "Rossi", "rossi@example.com" } };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");

            // Act - No sheet selected
            bool canSend = _controller.CanSendEmails();

            // Assert
            Assert.That(canSend, Is.False,
                "CanSendEmails should return false when no sheet is selected (Requirements 2.1, 5.1)");
        }

        /// <summary>
        /// Test CanSendEmails returns true when all prerequisites are met.
        /// **Validates: Requirements 2.1, 3.6, 5.1**
        /// </summary>
        [Test]
        public void CanSendEmails_WithAllPrerequisites_ShouldReturnTrue()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Configure Gmail credentials
            _controller.OnGmailCredentialsUpdated("test@gmail.com", "password");

            // Add volunteers
            var volunteers = new Dictionary<string, string> { { "Rossi", "rossi@example.com" } };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");

            // Select sheet
            _controller.OnSheetSelected("Sheet1");

            // Act
            bool canSend = _controller.CanSendEmails();

            // Assert
            Assert.That(canSend, Is.True,
                "CanSendEmails should return true when all prerequisites are met (Requirements 2.1, 3.6, 5.1)");
        }

        /// <summary>
        /// Test OnSheetSelected stores and persists the sheet name.
        /// **Validates: Requirement 2.4**
        /// </summary>
        [Test]
        public void OnSheetSelected_ShouldStoreAndPersistSheetName()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            var capturedConfig = new AppConfiguration();
            _mockConfigurationService
                .Setup(x => x.SaveConfiguration(It.IsAny<AppConfiguration>()))
                .Callback<AppConfiguration>(config => capturedConfig = config);

            // Act
            _controller.OnSheetSelected("TestSheet");

            // Assert
            _mockConfigurationService.Verify(
                x => x.SaveConfiguration(It.Is<AppConfiguration>(c => c.LastSheetName == "TestSheet")),
                Times.Once,
                "OnSheetSelected should persist the sheet name to configuration (Requirement 2.4)");

            Assert.That(capturedConfig.LastSheetName, Is.EqualTo("TestSheet"),
                "The persisted configuration should contain the selected sheet name");
        }

        /// <summary>
        /// Test OnSheetSelected updates CanSendEmails state.
        /// **Validates: Requirement 2.4, 5.1**
        /// </summary>
        [Test]
        public void OnSheetSelected_ShouldUpdateCanSendEmailsState()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Act
            _controller.OnSheetSelected("TestSheet");

            // Assert
            _mockUI.Verify(x => x.EnableSendEmailsButton(It.IsAny<bool>()), Times.Once,
                "OnSheetSelected should update the send emails button state");
        }

        /// <summary>
        /// Test OnDeleteAllVolunteers prompts for confirmation.
        /// **Validates: Requirements 8.7**
        /// </summary>
        [Test]
        public void OnDeleteAllVolunteers_ShouldPromptForConfirmation()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            _mockUI.Setup(x => x.ConfirmAction(It.IsAny<string>())).Returns(false);

            // Act
            _controller.OnDeleteAllVolunteers();

            // Assert
            _mockUI.Verify(
                x => x.ConfirmAction(It.Is<string>(msg => msg.Contains("sicuro"))),
                Times.Once,
                "OnDeleteAllVolunteers should prompt user for confirmation with Italian message (Requirement 8.7)");
        }

        /// <summary>
        /// Test OnDeleteAllVolunteers only deletes if confirmed.
        /// **Validates: Requirements 8.7, 8.8**
        /// </summary>
        [Test]
        public void OnDeleteAllVolunteers_WhenNotConfirmed_ShouldNotDelete()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Load some volunteers first
            var volunteers = new Dictionary<string, string>
            {
                { "Rossi", "rossi@example.com" },
                { "Bianchi", "bianchi@example.com" }
            };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");

            // Reset the mock to track only calls after this point
            _mockVolunteerManager.Invocations.Clear();

            // User does NOT confirm
            _mockUI.Setup(x => x.ConfirmAction(It.IsAny<string>())).Returns(false);

            // Act
            _controller.OnDeleteAllVolunteers();

            // Assert
            _mockVolunteerManager.Verify(
                x => x.SaveVolunteers(It.IsAny<string>(), It.IsAny<Dictionary<string, string>>()),
                Times.Never,
                "OnDeleteAllVolunteers should not save when user cancels confirmation (Requirement 8.8)");

            var currentVolunteers = _controller.GetVolunteers();
            Assert.That(currentVolunteers.Count, Is.EqualTo(2),
                "Volunteers should not be deleted when user cancels confirmation");
        }

        /// <summary>
        /// Test OnDeleteAllVolunteers deletes all volunteers when confirmed.
        /// **Validates: Requirements 8.7, 8.8**
        /// </summary>
        [Test]
        public void OnDeleteAllVolunteers_WhenConfirmed_ShouldDeleteAllVolunteers()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Load some volunteers first
            var volunteers = new Dictionary<string, string>
            {
                { "Rossi", "rossi@example.com" },
                { "Bianchi", "bianchi@example.com" }
            };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");

            // User confirms deletion
            _mockUI.Setup(x => x.ConfirmAction(It.IsAny<string>())).Returns(true);

            // Act
            _controller.OnDeleteAllVolunteers();

            // Assert
            _mockVolunteerManager.Verify(
                x => x.SaveVolunteers(It.IsAny<string>(), It.Is<Dictionary<string, string>>(d => d.Count == 0)),
                Times.Once,
                "OnDeleteAllVolunteers should save empty volunteers dictionary when confirmed (Requirement 8.8)");

            _mockUI.Verify(
                x => x.DisplayVolunteerList(It.Is<Dictionary<string, string>>(d => d.Count == 0)),
                Times.AtLeastOnce,
                "OnDeleteAllVolunteers should refresh UI with empty list when confirmed (Requirement 8.11)");

            var currentVolunteers = _controller.GetVolunteers();
            Assert.That(currentVolunteers.Count, Is.EqualTo(0),
                "All volunteers should be deleted when user confirms (Requirement 8.8)");
        }

        /// <summary>
        /// Test OnDeleteAllVolunteers updates CanSendEmails state after deletion.
        /// **Validates: Requirements 8.8, 5.1**
        /// </summary>
        [Test]
        public void OnDeleteAllVolunteers_WhenConfirmed_ShouldUpdateCanSendEmailsState()
        {
            // Arrange
            _controller = new VolunteerNotificationController(
                _mockVolunteerManager.Object,
                _mockEmailService.Object,
                _mockConfigurationService.Object,
                _mockExcelManager.Object,
                _mockUI.Object);

            // Setup initial state with all prerequisites met
            var volunteers = new Dictionary<string, string> { { "Rossi", "rossi@example.com" } };
            _mockVolunteerManager
                .Setup(x => x.LoadVolunteers(It.IsAny<string>()))
                .Returns(volunteers);
            _controller.OnVolunteerFileSelected("volunteers.json");
            _controller.OnGmailCredentialsUpdated("test@gmail.com", "password");
            _controller.OnSheetSelected("Sheet1");

            // User confirms deletion
            _mockUI.Setup(x => x.ConfirmAction(It.IsAny<string>())).Returns(true);

            // Act
            _controller.OnDeleteAllVolunteers();

            // Assert
            bool canSend = _controller.CanSendEmails();
            Assert.That(canSend, Is.False,
                "CanSendEmails should return false after deleting all volunteers (Requirement 5.1)");

            _mockUI.Verify(x => x.EnableSendEmailsButton(false), Times.AtLeastOnce,
                "Send emails button should be disabled after deleting all volunteers");
        }
    }
}
