using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI.Controls
{
    /// <summary>
    /// Custom TextBox with colored bottom border, focus/blur color change, and placeholder support.
    /// Validates: Requirements 5.1-5.5
    /// </summary>
    public class ModernTextBox : TextBox
    {
        private bool _isFocused;
        private string _placeholderText = string.Empty;

        // Colors
        private static readonly Color BorderNormal   = Color.FromArgb(0x00, 0x92, 0x46); // Verde #009246
        private static readonly Color BorderFocused  = Color.FromArgb(0xFA, 0xB9, 0x00); // Ambra #FAB900
        private static readonly Color PlaceholderColor = Color.FromArgb(0xAA, 0xAA, 0xAA);

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
            ForeColor = Color.FromArgb(0x39, 0x39, 0x39);
            Font = new Font("Segoe UI", 9F);
            Height = 28;
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.DoubleBuffer, true);
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

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Draw bottom border
            var borderColor = _isFocused ? BorderFocused : BorderNormal;
            using (var pen = new Pen(borderColor, 2))
            {
                e.Graphics.DrawLine(pen, 0, Height - 2, Width, Height - 2);
            }

            // Draw placeholder when empty and not focused
            if (!_isFocused && string.IsNullOrEmpty(Text) && !string.IsNullOrEmpty(_placeholderText))
            {
                using (var brush = new SolidBrush(PlaceholderColor))
                {
                    var rect = new Rectangle(1, 2, Width - 2, Height - 4);
                    e.Graphics.DrawString(_placeholderText, Font, brush, rect);
                }
            }
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            // WM_PAINT = 0x000F — trigger repaint to keep border visible
            if (m.Msg == 0x000F)
            {
                using var g = Graphics.FromHwnd(Handle);
                var borderColor = _isFocused ? BorderFocused : BorderNormal;
                using var pen = new Pen(borderColor, 2);
                g.DrawLine(pen, 0, Height - 2, Width, Height - 2);
            }
        }
    }
}
