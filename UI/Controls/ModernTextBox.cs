using System;
using System.ComponentModel;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI.Controls
{
    /// <summary>
    /// Custom TextBox with colored bottom border, focus/blur color change, and placeholder support.
    /// Validates: Requirements 5.1-5.5
    /// </summary>
    public class ModernTextBox : TextBox
    {
        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string? pszSubAppName, string? pszSubIdList);

        private bool _isFocused;
        private string _placeholderText = string.Empty;

        // Colors
        private static readonly Color BorderNormal   = Color.FromArgb(0xBE, 0xDF, 0xCA); // #bedfca
        private static readonly Color BorderFocused  = Color.FromArgb(0x06, 0x85, 0x34); // #068534
        private static readonly Color PlaceholderColor = Color.FromArgb(0x7C, 0xBF, 0x94); // #7cbf94

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public string PlaceholderText
        {
            get => _placeholderText;
            set
            {
                _placeholderText = value ?? string.Empty;
                Invalidate();
            }
        }

        public ModernTextBox()
        {
            BorderStyle = BorderStyle.None;
            BackColor = Color.White;
            ForeColor = Color.Black;
            Font = new Font("Segoe UI", 9F);
            Height = 28;

            // IMPORTANT: TextBox wraps a real native Win32 "Edit" control — it is not
            // self-drawn. Turning on ControlStyles.UserPaint/AllPaintingInWmPaint (as a
            // previous version of this control did) makes WinForms intercept WM_PAINT
            // itself instead of letting the native edit control paint its own text.
            // The overlay (border/placeholder) below is drawn via OnPaint AND the
            // WM_PAINT hook in WndProc, so with UserPaint enabled the native text
            // painting was being swallowed — this is the root cause of the typed text
            // being invisible (it rendered as blank/white regardless of ForeColor).
            // Do NOT set UserPaint/AllPaintingInWmPaint here; let the native control
            // paint text normally and only layer our extra drawing on top.
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);

            // Disable Windows visual-style theming on the native edit control.
            // Explorer theming (uxtheme) can silently override the control's text
            // color — especially under Windows dark mode — making typed text render
            // white regardless of the ForeColor property. Turning theming off here
            // forces the control to respect ForeColor/BackColor directly.
            try { SetWindowTheme(Handle, "", ""); } catch { /* not fatal if unsupported */ }

            // Defensive: reassert explicit colors too, in case the host/theme
            // resets them when the handle is (re)created.
            ForeColor = Color.Black;
            BackColor = Color.White;
        }

        protected override void OnGotFocus(EventArgs e)
        {
            _isFocused = true;
            Invalidate();
            base.OnGotFocus(e);
        }

        protected override void OnLostFocus(EventArgs e)
        {
            _isFocused = false;
            Invalidate();
            base.OnLostFocus(e);
        }

        protected override void OnTextChanged(EventArgs e)
        {
            Invalidate();
            base.OnTextChanged(e);
        }

        protected override void WndProc(ref Message m)
        {
            // Let the native Edit control handle WM_PAINT (and everything else) first,
            // so it renders the typed text itself using ForeColor/BackColor. We then
            // layer our custom border (and placeholder text, when empty/unfocused) on
            // top by drawing directly onto the control's HWND.
            base.WndProc(ref m);

            // WM_PAINT = 0x000F — draw overlay after the native control has painted.
            if (m.Msg == 0x000F)
            {
                using var g = Graphics.FromHwnd(Handle);

                var borderColor = _isFocused ? BorderFocused : BorderNormal;
                using (var pen = new Pen(borderColor, 2))
                {
                    g.DrawLine(pen, 0, Height - 2, Width, Height - 2);
                }

                if (!_isFocused && string.IsNullOrEmpty(Text) && !string.IsNullOrEmpty(_placeholderText))
                {
                    using var brush = new SolidBrush(PlaceholderColor);
                    var rect = new Rectangle(1, 2, Width - 2, Height - 4);
                    g.DrawString(_placeholderText, Font, brush, rect);
                }
            }
        }
    }
}
