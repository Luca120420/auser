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
    /// Bug condition exploration tests for MainForm resizability issue.
    /// These tests verify the expected behavior for window resizing on small screens.
    /// **Validates: Requirements 1.1, 1.3, 1.4, 2.1, 2.2, 2.3, 2.4**
    /// </summary>
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)]
    public class MainFormBugConditionTests
    {
        private Mock<IApplicationController> _mockController = null!;

        [SetUp]
        public void SetUp()
        {
            _mockController = new Mock<IApplicationController>();
        }

        /// <summary>
        /// Property 1: Fault Condition - Window Resizable on Small Screens
        /// 
        /// This test encodes the EXPECTED behavior after the fix.
        /// On UNFIXED code, this test MUST FAIL because:
        /// - FormBorderStyle is FixedDialog (should be Sizable)
        /// - MaximizeBox is false (should be true)
        /// - MinimumSize is not set (should be 600x400)
        /// 
        /// The test failure will surface counterexamples demonstrating the bug exists.
        /// After the fix is implemented, this same test should PASS.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3, 2.4**
        /// </summary>
        [Test]
        public void Property_WindowResizableOnSmallScreens_ExpectedBehavior()
        {
            // Arrange - Create form instance
            using (var form = new MainForm(_mockController.Object))
            {
                // Act - Check form properties after initialization
                var formBorderStyle = form.FormBorderStyle;
                var maximizeBox = form.MaximizeBox;
                var minimumSize = form.MinimumSize;

                // Assert - Expected behavior (will FAIL on unfixed code)
                // These assertions encode what the behavior SHOULD be after the fix
                
                // Requirement 2.1, 2.3: Window should be resizable
                Assert.That(formBorderStyle, Is.EqualTo(FormBorderStyle.Sizable),
                    $"FormBorderStyle should be Sizable to allow window resizing. " +
                    $"COUNTEREXAMPLE: FormBorderStyle is {formBorderStyle} instead of Sizable");
                
                // Requirement 2.4: Window should allow maximization
                Assert.That(maximizeBox, Is.True,
                    "MaximizeBox should be true to allow window maximization on small screens. " +
                    "COUNTEREXAMPLE: MaximizeBox is false");
                
                // Requirement 2.3: Window should have minimum size constraints
                Assert.That(minimumSize.Width, Is.EqualTo(600),
                    $"MinimumSize width should be 600 to prevent window from being too small. " +
                    $"COUNTEREXAMPLE: MinimumSize.Width is {minimumSize.Width}");
                
                Assert.That(minimumSize.Height, Is.EqualTo(400),
                    $"MinimumSize height should be 400 to prevent window from being too small. " +
                    $"COUNTEREXAMPLE: MinimumSize.Height is {minimumSize.Height}");
            }
        }

        /// <summary>
        /// Property-based test using FsCheck to verify window resizability across different screen sizes.
        /// This test generates random screen heights less than 1000px and verifies the window
        /// can be resized to fit those screens.
        /// 
        /// **Validates: Requirements 2.1, 2.2, 2.3, 2.4**
        /// </summary>
        [Test]
        public void Property_WindowResizable_ForAllSmallScreenSizes()
        {
            // Define the property: For all screen heights < 1000px, window should be resizable
            Prop.ForAll<int>(screenHeight =>
            {
                // Scope to small screens (400-999 pixels)
                if (screenHeight < 400 || screenHeight >= 1000)
                    return true; // Skip out-of-scope values

                using (var form = new MainForm(_mockController.Object))
                {
                    // The window should be resizable regardless of screen size
                    var isResizable = form.FormBorderStyle == FormBorderStyle.Sizable;
                    var canMaximize = form.MaximizeBox;
                    var hasMinimumSize = form.MinimumSize.Width > 0 && form.MinimumSize.Height > 0;

                    // All three conditions must be true for proper resizability
                    return isResizable && canMaximize && hasMinimumSize;
                }
            }).QuickCheckThrowOnFailure();
        }

        /// <summary>
        /// Scoped property test focusing on the concrete failing case.
        /// Tests the specific bug condition: window with 1000px height on screens < 1000px.
        /// 
        /// **Validates: Requirements 1.1, 1.3, 1.4, 2.1, 2.2, 2.3, 2.4**
        /// </summary>
        [Test]
        public void Property_ScopedBugCondition_WindowNotResizableOnSmallScreens()
        {
            // This test focuses on the exact bug scenario described in the requirements
            using (var form = new MainForm(_mockController.Object))
            {
                // Bug condition: Window is 1000px tall (after VolunteerPanel is added)
                // and screen is smaller than 1000px
                var windowHeight = form.Height; // Will be 1000 after initialization
                
                // Simulate small screen scenarios (768px, 800px, 900px)
                var smallScreenHeights = new[] { 768, 800, 900 };
                
                foreach (var screenHeight in smallScreenHeights)
                {
                    // On small screens, the bug manifests:
                    // - Window extends beyond screen (windowHeight > screenHeight)
                    // - User cannot resize window (FormBorderStyle is FixedDialog)
                    // - User cannot maximize window (MaximizeBox is false)
                    
                    var isBugCondition = screenHeight < windowHeight;
                    
                    if (isBugCondition)
                    {
                        // Expected behavior: Window should be resizable
                        Assert.That(form.FormBorderStyle, Is.EqualTo(FormBorderStyle.Sizable),
                            $"On {screenHeight}px screen with {windowHeight}px window: " +
                            $"FormBorderStyle should be Sizable. COUNTEREXAMPLE: {form.FormBorderStyle}");
                        
                        Assert.That(form.MaximizeBox, Is.True,
                            $"On {screenHeight}px screen with {windowHeight}px window: " +
                            $"MaximizeBox should be true. COUNTEREXAMPLE: MaximizeBox is false");
                        
                        Assert.That(form.MinimumSize, Is.Not.EqualTo(Size.Empty),
                            $"On {screenHeight}px screen with {windowHeight}px window: " +
                            $"MinimumSize should be set. COUNTEREXAMPLE: MinimumSize is {form.MinimumSize}");
                    }
                }
            }
        }
    }
}
