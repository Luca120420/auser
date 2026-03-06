using NUnit.Framework;
using Moq;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Services;
using System.Windows.Forms;
using System.Drawing;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for MainForm responsive layout behavior.
    /// Validates: Requirements 1.1, 1.2, 1.3, 1.4
    /// </summary>
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)] // Required for Windows Forms
    public class MainFormResponsiveLayoutTests
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

        /// <summary>
        /// Test that all 4 buttons have Top|Left anchoring after initialization.
        /// Validates: Requirements 3.1, 3.2, 3.3, 3.4
        /// </summary>
        [Test]
        public void ButtonsHaveCorrectAnchoring()
        {
            // Arrange - Form is created in SetUp
            var expectedAnchor = AnchorStyles.Top | AnchorStyles.Left;

            // Act - Find all buttons
            Button? btnSelectCSV = null;
            Button? btnSelectExcel = null;
            Button? btnProcess = null;
            Button? btnDownload = null;

            foreach (Control control in _form.Controls)
            {
                if (control is Button button)
                {
                    if (button.Text.Contains("CSV") || button.Text.Contains("csv"))
                        btnSelectCSV = button;
                    else if (button.Text.Contains("Excel") || button.Text.Contains("excel"))
                        btnSelectExcel = button;
                    else if (button.Text.Contains("Process") || button.Text.Contains("process") || button.Text.Contains("Elabora"))
                        btnProcess = button;
                    else if (button.Text.Contains("Download") || button.Text.Contains("download") || button.Text.Contains("Scarica"))
                        btnDownload = button;
                }
            }

            // Assert - Verify all buttons exist and have correct anchoring
            Assert.That(btnSelectCSV, Is.Not.Null, "CSV selection button should exist");
            Assert.That(btnSelectCSV!.Anchor, Is.EqualTo(expectedAnchor), 
                $"CSV button anchoring should be Top|Left. Found: {btnSelectCSV.Anchor}");

            Assert.That(btnSelectExcel, Is.Not.Null, "Excel selection button should exist");
            Assert.That(btnSelectExcel!.Anchor, Is.EqualTo(expectedAnchor), 
                $"Excel button anchoring should be Top|Left. Found: {btnSelectExcel.Anchor}");

            Assert.That(btnProcess, Is.Not.Null, "Process button should exist");
            Assert.That(btnProcess!.Anchor, Is.EqualTo(expectedAnchor), 
                $"Process button anchoring should be Top|Left. Found: {btnProcess.Anchor}");

            Assert.That(btnDownload, Is.Not.Null, "Download button should exist");
            Assert.That(btnDownload!.Anchor, Is.EqualTo(expectedAnchor), 
                $"Download button anchoring should be Top|Left. Found: {btnDownload.Anchor}");
        }

        /// <summary>
        /// Test that all 3 expanding labels have Top|Left|Right anchoring after initialization.
        /// Validates: Requirements 3.5, 3.6, 3.7
        /// </summary>
        [Test]
        public void ExpandingLabelsHaveCorrectAnchoring()
        {
            // Arrange - Form is created in SetUp
            var expectedAnchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;

            // Act - Find all expanding labels (path labels and status label)
            Label? lblCSVPath = null;
            Label? lblExcelPath = null;
            Label? lblStatus = null;

            foreach (Control control in _form.Controls)
            {
                if (control is Label label)
                {
                    // Path labels have ForeColor = DarkBlue and are positioned at X=120
                    if (label.ForeColor == Color.DarkBlue && label.Location.X == 120)
                    {
                        if (label.Location.Y == 70)
                            lblCSVPath = label;
                        else if (label.Location.Y == 150)
                            lblExcelPath = label;
                    }
                    // Status label is at Y=250
                    else if (label.Location.Y == 250)
                    {
                        lblStatus = label;
                    }
                }
            }

            // Assert - Verify all expanding labels exist and have correct anchoring
            Assert.That(lblCSVPath, Is.Not.Null, "CSV path label should exist");
            Assert.That(lblCSVPath!.Anchor, Is.EqualTo(expectedAnchor), 
                $"CSV path label anchoring should be Top|Left|Right. Found: {lblCSVPath.Anchor}");

            Assert.That(lblExcelPath, Is.Not.Null, "Excel path label should exist");
            Assert.That(lblExcelPath!.Anchor, Is.EqualTo(expectedAnchor), 
                $"Excel path label anchoring should be Top|Left|Right. Found: {lblExcelPath.Anchor}");

            Assert.That(lblStatus, Is.Not.Null, "Status label should exist");
            Assert.That(lblStatus!.Anchor, Is.EqualTo(expectedAnchor), 
                $"Status label anchoring should be Top|Left|Right. Found: {lblStatus.Anchor}");
        }

        /// <summary>
        /// Test that VolunteerPanel retains Top|Left|Right|Bottom anchoring after initialization.
        /// Validates: Requirement 1.4
        /// </summary>
        [Test]
        public void VolunteerPanelRetainsCorrectAnchoring()
        {
            // Arrange - Form is created in SetUp
            var expectedAnchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;

            // Act - Find the VolunteerPanel control
            VolunteerPanel? volunteerPanel = null;

            foreach (Control control in _form.Controls)
            {
                if (control is VolunteerPanel panel)
                {
                    volunteerPanel = panel;
                    break;
                }
            }

            // Assert - Verify VolunteerPanel exists and has correct anchoring
            Assert.That(volunteerPanel, Is.Not.Null, "VolunteerPanel should exist");
            Assert.That(volunteerPanel!.Anchor, Is.EqualTo(expectedAnchor), 
                $"VolunteerPanel anchoring should be Top|Left|Right|Bottom. Found: {volunteerPanel.Anchor}");
        }

        /// <summary>
        /// Test that MinimumSize property is set to (600, 400).
        /// Validates: Requirements 2.1, 2.2
        /// </summary>
        [Test]
        public void MinimumSizeIsSetCorrectly()
        {
            // Arrange - Form is created in SetUp
            var expectedMinimumSize = new Size(600, 400);

            // Act - Get the MinimumSize property
            var actualMinimumSize = _form.MinimumSize;

            // Assert - Verify MinimumSize is set to (600, 400)
            Assert.That(actualMinimumSize, Is.EqualTo(expectedMinimumSize), 
                $"MainForm MinimumSize should be (600, 400). Found: {actualMinimumSize}");
        }

        /// <summary>
        /// Test that attempting to resize below minimum is prevented by framework.
        /// Validates: Requirement 2.3
        /// </summary>
        [Test]
        public void ResizeBelowMinimumIsPrevented()
        {
            // Arrange - Form is created in SetUp
            var minimumSize = new Size(600, 400);
            var belowMinimumSize = new Size(500, 300);

            // Act - Attempt to resize below minimum
            _form.Size = belowMinimumSize;

            // Assert - Verify the form size is clamped to minimum
            Assert.That(_form.Width, Is.GreaterThanOrEqualTo(minimumSize.Width), 
                $"Form width should not go below {minimumSize.Width}. Found: {_form.Width}");
            Assert.That(_form.Height, Is.GreaterThanOrEqualTo(minimumSize.Height), 
                $"Form height should not go below {minimumSize.Height}. Found: {_form.Height}");
        }

        /// <summary>
        /// Test form size at boundary condition (601x401) - just above minimum.
        /// Validates: Requirements 2.1, 2.2, 2.3
        /// </summary>
        [Test]
        public void FormSizeAtBoundaryCondition()
        {
            // Arrange - Form is created in SetUp
            var boundarySize = new Size(601, 401);

            // Act - Set form size to just above minimum
            _form.Size = boundarySize;

            // Assert - Verify the form accepts this size (it's above minimum)
            Assert.That(_form.Width, Is.EqualTo(boundarySize.Width), 
                $"Form should accept width of {boundarySize.Width}. Found: {_form.Width}");
            Assert.That(_form.Height, Is.EqualTo(boundarySize.Height), 
                $"Form should accept height of {boundarySize.Height}. Found: {_form.Height}");
        }

        /// <summary>
        /// Test resize from 600x400 to 800x600 - button positions unchanged, labels expanded.
        /// Validates: Requirements 1.1, 1.2, 4.1, 4.4
        /// </summary>
        [Test]
        public void ResizeFrom600x400To800x600()
        {
            // Arrange - Set form to minimum size
            _form.Size = new Size(600, 400);

            // Find controls
            Button? btnSelectCSV = null;
            Button? btnProcess = null;
            Label? lblCSVPath = null;
            Label? lblStatus = null;

            foreach (Control control in _form.Controls)
            {
                if (control is Button button)
                {
                    if (button.Text.Contains("CSV") || button.Text.Contains("csv"))
                        btnSelectCSV = button;
                    else if (button.Text.Contains("Process") || button.Text.Contains("process") || button.Text.Contains("Elabora"))
                        btnProcess = button;
                }
                else if (control is Label label)
                {
                    if (label.ForeColor == Color.DarkBlue && label.Location.X == 120 && label.Location.Y == 70)
                        lblCSVPath = label;
                    else if (label.Location.Y == 250)
                        lblStatus = label;
                }
            }

            Assert.That(btnSelectCSV, Is.Not.Null, "CSV button should exist");
            Assert.That(btnProcess, Is.Not.Null, "Process button should exist");
            Assert.That(lblCSVPath, Is.Not.Null, "CSV path label should exist");
            Assert.That(lblStatus, Is.Not.Null, "Status label should exist");

            // Record initial positions and sizes
            int initialCSVButtonLeft = btnSelectCSV!.Left;
            int initialProcessButtonLeft = btnProcess!.Left;
            int initialCSVPathWidth = lblCSVPath!.Width;
            int initialStatusWidth = lblStatus!.Width;

            // Act - Resize to 800x600
            _form.Size = new Size(800, 600);

            // Assert - Button positions unchanged
            Assert.That(btnSelectCSV.Left, Is.EqualTo(initialCSVButtonLeft), 
                "CSV button Left position should remain unchanged after resize");
            Assert.That(btnProcess.Left, Is.EqualTo(initialProcessButtonLeft), 
                "Process button Left position should remain unchanged after resize");

            // Assert - Labels expanded by 200 pixels (800 - 600)
            int expectedWidthIncrease = 200;
            Assert.That(lblCSVPath.Width, Is.EqualTo(initialCSVPathWidth + expectedWidthIncrease), 
                $"CSV path label should expand by {expectedWidthIncrease} pixels. Initial: {initialCSVPathWidth}, Expected: {initialCSVPathWidth + expectedWidthIncrease}, Found: {lblCSVPath.Width}");
            Assert.That(lblStatus.Width, Is.EqualTo(initialStatusWidth + expectedWidthIncrease), 
                $"Status label should expand by {expectedWidthIncrease} pixels. Initial: {initialStatusWidth}, Expected: {initialStatusWidth + expectedWidthIncrease}, Found: {lblStatus.Width}");
        }

        /// <summary>
        /// Test resize from 800x600 to 1024x768 - verify proportional label expansion.
        /// Validates: Requirements 1.1, 1.2, 4.1, 4.4
        /// </summary>
        [Test]
        public void ResizeFrom800x600To1024x768()
        {
            // Arrange - Set form to 800x600
            _form.Size = new Size(800, 600);

            // Find controls
            Button? btnSelectExcel = null;
            Button? btnDownload = null;
            Label? lblExcelPath = null;
            Label? lblStatus = null;

            foreach (Control control in _form.Controls)
            {
                if (control is Button button)
                {
                    if (button.Text.Contains("Excel") || button.Text.Contains("excel"))
                        btnSelectExcel = button;
                    else if (button.Text.Contains("Download") || button.Text.Contains("download") || button.Text.Contains("Scarica"))
                        btnDownload = button;
                }
                else if (control is Label label)
                {
                    if (label.ForeColor == Color.DarkBlue && label.Location.X == 120 && label.Location.Y == 150)
                        lblExcelPath = label;
                    else if (label.Location.Y == 250)
                        lblStatus = label;
                }
            }

            Assert.That(btnSelectExcel, Is.Not.Null, "Excel button should exist");
            Assert.That(btnDownload, Is.Not.Null, "Download button should exist");
            Assert.That(lblExcelPath, Is.Not.Null, "Excel path label should exist");
            Assert.That(lblStatus, Is.Not.Null, "Status label should exist");

            // Record initial positions and sizes
            int initialExcelButtonLeft = btnSelectExcel!.Left;
            int initialDownloadButtonLeft = btnDownload!.Left;
            int initialExcelPathWidth = lblExcelPath!.Width;
            int initialStatusWidth = lblStatus!.Width;

            // Act - Resize to 1024x768
            _form.Size = new Size(1024, 768);

            // Assert - Button positions unchanged
            Assert.That(btnSelectExcel.Left, Is.EqualTo(initialExcelButtonLeft), 
                "Excel button Left position should remain unchanged after resize");
            Assert.That(btnDownload.Left, Is.EqualTo(initialDownloadButtonLeft), 
                "Download button Left position should remain unchanged after resize");

            // Assert - Labels expanded by 224 pixels (1024 - 800)
            int expectedWidthIncrease = 224;
            Assert.That(lblExcelPath.Width, Is.EqualTo(initialExcelPathWidth + expectedWidthIncrease), 
                $"Excel path label should expand by {expectedWidthIncrease} pixels. Initial: {initialExcelPathWidth}, Expected: {initialExcelPathWidth + expectedWidthIncrease}, Found: {lblExcelPath.Width}");
            Assert.That(lblStatus.Width, Is.EqualTo(initialStatusWidth + expectedWidthIncrease), 
                $"Status label should expand by {expectedWidthIncrease} pixels. Initial: {initialStatusWidth}, Expected: {initialStatusWidth + expectedWidthIncrease}, Found: {lblStatus.Width}");
        }

        /// <summary>
        /// Test resize from 1024x768 back to 600x400 - verify controls return to original state.
        /// Validates: Requirements 1.1, 1.2, 4.1, 4.4
        /// </summary>
        [Test]
        public void ResizeFrom1024x768BackTo600x400()
        {
            // Arrange - Set form to minimum size first to record original state
            _form.Size = new Size(600, 400);

            // Find controls
            Button? btnSelectCSV = null;
            Button? btnSelectExcel = null;
            Button? btnProcess = null;
            Button? btnDownload = null;
            Label? lblCSVPath = null;
            Label? lblExcelPath = null;
            Label? lblStatus = null;

            foreach (Control control in _form.Controls)
            {
                if (control is Button button)
                {
                    if (button.Text.Contains("CSV") || button.Text.Contains("csv"))
                        btnSelectCSV = button;
                    else if (button.Text.Contains("Excel") || button.Text.Contains("excel"))
                        btnSelectExcel = button;
                    else if (button.Text.Contains("Process") || button.Text.Contains("process") || button.Text.Contains("Elabora"))
                        btnProcess = button;
                    else if (button.Text.Contains("Download") || button.Text.Contains("download") || button.Text.Contains("Scarica"))
                        btnDownload = button;
                }
                else if (control is Label label)
                {
                    if (label.ForeColor == Color.DarkBlue && label.Location.X == 120)
                    {
                        if (label.Location.Y == 70)
                            lblCSVPath = label;
                        else if (label.Location.Y == 150)
                            lblExcelPath = label;
                    }
                    else if (label.Location.Y == 250)
                        lblStatus = label;
                }
            }

            Assert.That(btnSelectCSV, Is.Not.Null, "CSV button should exist");
            Assert.That(btnSelectExcel, Is.Not.Null, "Excel button should exist");
            Assert.That(btnProcess, Is.Not.Null, "Process button should exist");
            Assert.That(btnDownload, Is.Not.Null, "Download button should exist");
            Assert.That(lblCSVPath, Is.Not.Null, "CSV path label should exist");
            Assert.That(lblExcelPath, Is.Not.Null, "Excel path label should exist");
            Assert.That(lblStatus, Is.Not.Null, "Status label should exist");

            // Record original state at 600x400
            int originalCSVButtonLeft = btnSelectCSV!.Left;
            int originalExcelButtonLeft = btnSelectExcel!.Left;
            int originalProcessButtonLeft = btnProcess!.Left;
            int originalDownloadButtonLeft = btnDownload!.Left;
            int originalCSVPathWidth = lblCSVPath!.Width;
            int originalExcelPathWidth = lblExcelPath!.Width;
            int originalStatusWidth = lblStatus!.Width;

            // Act - Resize to 1024x768 and then back to 600x400
            _form.Size = new Size(1024, 768);
            _form.Size = new Size(600, 400);

            // Assert - All button positions returned to original state
            Assert.That(btnSelectCSV.Left, Is.EqualTo(originalCSVButtonLeft), 
                "CSV button Left position should return to original state");
            Assert.That(btnSelectExcel.Left, Is.EqualTo(originalExcelButtonLeft), 
                "Excel button Left position should return to original state");
            Assert.That(btnProcess.Left, Is.EqualTo(originalProcessButtonLeft), 
                "Process button Left position should return to original state");
            Assert.That(btnDownload.Left, Is.EqualTo(originalDownloadButtonLeft), 
                "Download button Left position should return to original state");

            // Assert - All label widths returned to original state
            Assert.That(lblCSVPath.Width, Is.EqualTo(originalCSVPathWidth), 
                $"CSV path label width should return to original state. Expected: {originalCSVPathWidth}, Found: {lblCSVPath.Width}");
            Assert.That(lblExcelPath.Width, Is.EqualTo(originalExcelPathWidth), 
                $"Excel path label width should return to original state. Expected: {originalExcelPathWidth}, Found: {lblExcelPath.Width}");
            Assert.That(lblStatus.Width, Is.EqualTo(originalStatusWidth), 
                $"Status label width should return to original state. Expected: {originalStatusWidth}, Found: {lblStatus.Width}");
        }
    }
}
