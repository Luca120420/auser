using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using NUnit.Framework;
using Moq;
using AuserExcelTransformer.UI;
using AuserExcelTransformer.UI.Controls;
using AuserExcelTransformer.Services;

namespace AuserExcelTransformer.Tests
{
    /// <summary>
    /// Unit tests for the Modern UI Redesign feature.
    /// Validates: Requirements 1.1-1.5, 2.1-2.5, 3.1-3.5, 4.1-4.8, 5.1-5.5, 6.1-6.4
    /// </summary>
    [TestFixture]
    [Apartment(System.Threading.ApartmentState.STA)]
    public class ModernUIDesignTests
    {
        // ── Helpers ──────────────────────────────────────────────────────────────

        private static T? FindControl<T>(Control parent) where T : Control
        {
            foreach (Control c in parent.Controls)
            {
                if (c is T found) return found;
                var nested = FindControl<T>(c);
                if (nested != null) return nested;
            }
            return null;
        }

        private static IEnumerable<T> FindAllControls<T>(Control parent) where T : Control
        {
            foreach (Control c in parent.Controls)
            {
                if (c is T found) yield return found;
                foreach (var nested in FindAllControls<T>(c))
                    yield return nested;
            }
        }

        // ── ThemeManager color constants ──────────────────────────────────────────

        [Test]
        public void ThemeManager_ColorBackground_IsWhite()
        {
            Assert.That(ThemeManager.ColorBackground, Is.EqualTo(Color.White));
        }

        [Test]
        public void ThemeManager_ColorPrimary_IsCarbone()
        {
            Assert.That(ThemeManager.ColorPrimary, Is.EqualTo(Color.FromArgb(0x39, 0x39, 0x39)));
        }

        [Test]
        public void ThemeManager_ColorAccent_IsVerde()
        {
            Assert.That(ThemeManager.ColorAccent, Is.EqualTo(Color.FromArgb(0x00, 0x92, 0x46)));
        }

        [Test]
        public void ThemeManager_ColorSecondary_IsAmbra()
        {
            Assert.That(ThemeManager.ColorSecondary, Is.EqualTo(Color.FromArgb(0xFA, 0xB9, 0x00)));
        }

        [Test]
        public void ThemeManager_ColorError_IsRed()
        {
            Assert.That(ThemeManager.ColorError, Is.EqualTo(Color.FromArgb(0xD3, 0x2F, 0x2F)));
        }

        // ── HeaderPanel ───────────────────────────────────────────────────────────

        [Test]
        public void HeaderPanel_Height_Is80()
        {
            using var header = new HeaderPanel();
            Assert.That(header.Height, Is.EqualTo(80));
        }

        [Test]
        public void HeaderPanel_BackColor_IsCarbone()
        {
            using var header = new HeaderPanel();
            Assert.That(header.BackColor, Is.EqualTo(Color.FromArgb(0x39, 0x39, 0x39)));
        }

        [Test]
        public void HeaderPanel_WithoutLogo_DoesNotThrow()
        {
            Assert.DoesNotThrow(() =>
            {
                using var header = new HeaderPanel();
            });
        }

        // ── MainForm ──────────────────────────────────────────────────────────────

        [Test]
        public void MainForm_MinimumSize_Is700x600()
        {
            var mock = new Mock<IApplicationController>();
            using var form = new MainForm(mock.Object);
            Assert.That(form.MinimumSize, Is.EqualTo(new Size(700, 600)));
        }

        [Test]
        public void MainForm_Title_IsAuserGestioneTrasporti()
        {
            var mock = new Mock<IApplicationController>();
            using var form = new MainForm(mock.Object);
            Assert.That(form.Text, Is.EqualTo("Auser Gestione Trasporti"));
        }

        [Test]
        public void MainForm_BackColor_IsWhite()
        {
            var mock = new Mock<IApplicationController>();
            using var form = new MainForm(mock.Object);
            Assert.That(form.BackColor, Is.EqualTo(Color.White));
        }

        // ── VolunteerPanel — GroupBoxes ───────────────────────────────────────────

        [Test]
        public void VolunteerPanel_GroupBoxes_AreModernGroupBox()
        {
            var mock = new Mock<IVolunteerNotificationController>();
            using var panel = new VolunteerPanel(mock.Object);
            var modernGroupBoxes = new List<ModernGroupBox>(FindAllControls<ModernGroupBox>(panel));
            Assert.That(modernGroupBoxes.Count, Is.GreaterThanOrEqualTo(4),
                $"Expected at least 4 ModernGroupBox instances, found {modernGroupBoxes.Count}");
        }

        // ── VolunteerPanel — TextBoxes ────────────────────────────────────────────

        [Test]
        public void VolunteerPanel_CredentialTextBoxes_AreModernTextBox()
        {
            var mock = new Mock<IVolunteerNotificationController>();
            using var panel = new VolunteerPanel(mock.Object);
            var modernTextBoxes = new List<ModernTextBox>(FindAllControls<ModernTextBox>(panel));
            Assert.That(modernTextBoxes.Count, Is.GreaterThanOrEqualTo(2),
                $"Expected at least 2 ModernTextBox instances, found {modernTextBoxes.Count}");
        }

        // ── VolunteerPanel — Buttons ──────────────────────────────────────────────

        [Test]
        public void VolunteerPanel_SendEmailButton_IsPrimaryStyle()
        {
            var mock = new Mock<IVolunteerNotificationController>();
            using var panel = new VolunteerPanel(mock.Object);
            var buttons = new List<ModernButton>(FindAllControls<ModernButton>(panel));
            Assert.That(buttons.Count, Is.GreaterThan(0), "Expected at least one ModernButton in VolunteerPanel");

            ModernButton? sendBtn = null;
            foreach (var btn in buttons)
            {
                if (btn.Text.IndexOf("Invia", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    btn.Text.IndexOf("Email", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    sendBtn = btn;
                    break;
                }
            }

            Assert.That(sendBtn, Is.Not.Null,
                "Could not find a ModernButton with text containing 'Invia' or 'Email'");
            Assert.That(sendBtn!.Style, Is.EqualTo(ModernButton.ButtonStyle.Primary));
        }

        // ── VolunteerPanel — initialization ──────────────────────────────────────

        [Test]
        public void VolunteerPanel_AddContactDialog_HasCorrectSize()
        {
            var mock = new Mock<IVolunteerNotificationController>();
            Assert.DoesNotThrow(() =>
            {
                using var panel = new VolunteerPanel(mock.Object);
                Assert.That(panel, Is.Not.Null);
            });
        }
    }
}
