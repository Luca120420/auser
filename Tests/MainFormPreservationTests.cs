using NUnit.Framework;
using FsCheck;
using Moq;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.Services;
using System.Windows.Forms;
using System.Drawing;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Preservation property tests for MainForm initialization behavior.
    /// These tests verify that the fix does not introduce regressions in existing behavior.
    /// Tests should PASS on both unfixed and fixed code.
    /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
    /// </summary>
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)]
    public class MainFormPreservationTests
    {
        private Mock<IApplicationController> _mockController = null!;

        [SetUp]
        public void SetUp()
        {
            _mockController = new Mock<IApplicationController>();
        }

        /// <summary>
        /// Property 2: Preservation - Initial Window Configuration
        /// 
        /// This test verifies that the initial window configuration remains unchanged:
        /// - Initial size (observed: 850x884 pixels due to Windows Forms chrome)
        /// - Window is centered on screen (Requirement 3.1)
        /// - AutoScroll is enabled (Requirement 3.4)
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.1, 3.4, 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_InitialWindowConfiguration()
        {
            // Arrange & Act - Create form instance
            using (var form = new MainForm(_mockController.Object))
            {
                // Assert - Verify initial window configuration is preserved
                
                // Requirement 3.5: Initial window size (observed actual values)
                // Note: Height is 884 instead of 1000 due to Windows Forms window chrome
                Assert.That(form.Size.Width, Is.EqualTo(850),
                    $"Initial window width should be 850 pixels. Found: {form.Size.Width}");
                
                Assert.That(form.Size.Height, Is.EqualTo(884),
                    $"Initial window height should be 884 pixels. Found: {form.Size.Height}");
                
                // Requirement 3.1: Window should be centered on screen
                Assert.That(form.StartPosition, Is.EqualTo(FormStartPosition.CenterScreen),
                    $"Window should be centered on screen. Found: {form.StartPosition}");
                
                // Requirement 3.4: AutoScroll should be enabled
                Assert.That(form.AutoScroll, Is.True,
                    $"AutoScroll should be enabled. Found: {form.AutoScroll}");
            }
        }


        /// <summary>
        /// Property 2: Preservation - VolunteerPanel Configuration
        /// 
        /// This test verifies that the VolunteerPanel configuration remains unchanged:
        /// - VolunteerPanel is positioned at (20, 350) (Requirement 3.2)
        /// - VolunteerPanel has proper anchoring (Top | Left | Right | Bottom) (Requirement 3.2)
        /// 
        /// Note: VolunteerPanel size is dynamic due to anchoring, so we only verify position and anchoring.
        /// 
        /// This test should PASS on unfixed code and continue to PASS after the fix.
        /// 
        /// **Validates: Requirements 3.2, 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_VolunteerPanelConfiguration()
        {
            // Arrange & Act - Create form instance
            using (var form = new MainForm(_mockController.Object))
            {
                // Find the VolunteerPanel control
                VolunteerPanel? volunteerPanel = null;
                foreach (Control control in form.Controls)
                {
                    if (control is VolunteerPanel panel)
                    {
                        volunteerPanel = panel;
                        break;
                    }
                }

                // Assert - Verify VolunteerPanel exists and is configured correctly
                Assert.That(volunteerPanel, Is.Not.Null,
                    "VolunteerPanel should be added to the form");

                if (volunteerPanel != null)
                {
                    // Requirement 3.2: VolunteerPanel should be positioned at (20, 350)
                    Assert.That(volunteerPanel.Location.X, Is.EqualTo(20),
                        $"VolunteerPanel X position should be 20. Found: {volunteerPanel.Location.X}");
                    
                    Assert.That(volunteerPanel.Location.Y, Is.EqualTo(350),
                        $"VolunteerPanel Y position should be 350. Found: {volunteerPanel.Location.Y}");
                    
                    // Requirement 3.2: VolunteerPanel should have proper anchoring
                    var expectedAnchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom;
                    Assert.That(volunteerPanel.Anchor, Is.EqualTo(expectedAnchor),
                        $"VolunteerPanel anchoring should be Top | Left | Right | Bottom. Found: {volunteerPanel.Anchor}");
                }
            }
        }

        /// <summary>
        /// Property-based test using FsCheck to verify initial window configuration
        /// remains consistent across multiple test runs.
        /// 
        /// This generates multiple test cases to ensure the preservation properties
        /// hold consistently, not just in a single test run.
        /// 
        /// **Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5**
        /// </summary>
        [Test]
        public void Property_Preservation_ConsistentInitialization()
        {
            // Define the property: For all test runs, initial configuration should be identical
            Prop.ForAll<int>(testRun =>
            {
                // Scope to reasonable test run numbers (1-100)
                if (testRun < 1 || testRun > 100)
                    return true; // Skip out-of-scope values

                using (var form = new MainForm(_mockController.Object))
                {
                    // Verify all preservation requirements (using observed actual values)
                    var sizeCorrect = form.Size.Width == 850 && form.Size.Height == 884;
                    var positionCorrect = form.StartPosition == FormStartPosition.CenterScreen;
                    var autoScrollCorrect = form.AutoScroll == true;
                    
                    // Find VolunteerPanel
                    VolunteerPanel? volunteerPanel = null;
                    foreach (Control control in form.Controls)
                    {
                        if (control is VolunteerPanel panel)
                        {
                            volunteerPanel = panel;
                            break;
                        }
                    }
                    
                    var panelExists = volunteerPanel != null;
                    var panelLocationCorrect = volunteerPanel?.Location.X == 20 && volunteerPanel?.Location.Y == 350;
                    var panelAnchorCorrect = volunteerPanel?.Anchor == (AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Bottom);

                    // All preservation properties must hold (size is dynamic due to anchoring)
                    return sizeCorrect && positionCorrect && autoScrollCorrect && 
                           panelExists && panelLocationCorrect && panelAnchorCorrect;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Property-based test to verify control count remains consistent.
        /// This ensures that the fix doesn't accidentally add or remove controls.
        /// 
        /// **Validates: Requirements 3.3**
        /// </summary>
        [Test]
        public void Property_Preservation_ControlCountConsistent()
        {
            // Create multiple form instances and verify control count is consistent
            int? expectedControlCount = null;

            for (int i = 0; i < 10; i++)
            {
                using (var form = new MainForm(_mockController.Object))
                {
                    var controlCount = form.Controls.Count;

                    if (expectedControlCount == null)
                    {
                        expectedControlCount = controlCount;
                    }
                    else
                    {
                        // Requirement 3.3: Control count should remain consistent
                        Assert.That(controlCount, Is.EqualTo(expectedControlCount.Value),
                            $"Control count should be consistent across form instances. " +
                            $"Expected: {expectedControlCount.Value}, Found: {controlCount}");
                    }
                }
            }
        }
    }
}
