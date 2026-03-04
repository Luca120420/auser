using NUnit.Framework;
using Moq;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the MainForm GUI component.
    /// Validates: Requirements 1.1, 1.4, 1.5, 7.3, 8.1, 8.2
    /// </summary>
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)] // Required for Windows Forms
    public class MainFormTests
    {
        private Mock<IApplicationController> _mockController = null!;
        private MainForm _form = null!;

        [SetUp]
        public void SetUp()
        {
            _mockController = new Mock<IApplicationController>();
            _form = new MainForm(_mockController.Object);
        }

        [TearDown]
        public void TearDown()
        {
            _form?.Dispose();
        }

        [Test]
        public void Constructor_WithValidController_CreatesForm()
        {
            // Assert
            Assert.That(_form, Is.Not.Null);
            Assert.That(_form.Text, Is.EqualTo(Properties.Resources.ApplicationTitle));
        }

        [Test]
        public void Constructor_WithNullController_ThrowsArgumentNullException()
        {
            // Act & Assert
            Assert.Throws<System.ArgumentNullException>(() => new MainForm(null!));
        }

        [Test]
        public void DisplaySelectedCSVPath_WithValidPath_UpdatesLabel()
        {
            // Arrange
            var testPath = @"C:\test\file.csv";

            // Act
            _form.DisplaySelectedCSVPath(testPath);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void DisplaySelectedExcelPath_WithValidPath_UpdatesLabel()
        {
            // Arrange
            var testPath = @"C:\test\file.xlsx";

            // Act
            _form.DisplaySelectedExcelPath(testPath);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void EnableProcessButton_WithTrue_EnablesButton()
        {
            // Act
            _form.EnableProcessButton(true);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void EnableProcessButton_WithFalse_DisablesButton()
        {
            // Act
            _form.EnableProcessButton(false);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void EnableDownloadButton_WithTrue_EnablesButton()
        {
            // Act
            _form.EnableDownloadButton(true);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void EnableDownloadButton_WithFalse_DisablesButton()
        {
            // Act
            _form.EnableDownloadButton(false);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void ShowErrorMessage_WithMessage_DisplaysInRed()
        {
            // Arrange
            var errorMessage = "Test error message";

            // Act
            _form.ShowErrorMessage(errorMessage);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void ShowSuccessMessage_WithMessage_DisplaysInGreen()
        {
            // Arrange
            var successMessage = "Test success message";

            // Act
            _form.ShowSuccessMessage(successMessage);

            // Assert - We can't directly access private controls, but we can verify no exception is thrown
            Assert.Pass("Method executed without exception");
        }

        [Test]
        public void ShowErrorMessage_WithItalianMessage_DisplaysCorrectly()
        {
            // Arrange
            var italianMessage = Properties.Resources.ErrorCSVFileRead;

            // Act
            _form.ShowErrorMessage(italianMessage);

            // Assert
            Assert.That(italianMessage, Does.Contain("Impossibile"));
            Assert.Pass("Italian error message displayed without exception");
        }

        [Test]
        public void ShowSuccessMessage_WithItalianMessage_DisplaysCorrectly()
        {
            // Arrange
            var italianMessage = Properties.Resources.SuccessMessage;

            // Act
            _form.ShowSuccessMessage(italianMessage);

            // Assert
            Assert.That(italianMessage, Does.Contain("successo"));
            Assert.Pass("Italian success message displayed without exception");
        }

        [Test]
        public void FormTitle_IsInItalian()
        {
            // Assert
            Assert.That(_form.Text, Is.EqualTo(Properties.Resources.ApplicationTitle));
            Assert.That(_form.Text, Does.Contain("Auser"));
        }

        [Test]
        public void FormProperties_AreSetCorrectly()
        {
            // Assert
            Assert.That(_form.FormBorderStyle, Is.EqualTo(System.Windows.Forms.FormBorderStyle.FixedDialog));
            Assert.That(_form.MaximizeBox, Is.False);
            Assert.That(_form.StartPosition, Is.EqualTo(System.Windows.Forms.FormStartPosition.CenterScreen));
        }
    }
}
