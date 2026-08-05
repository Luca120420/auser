using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI.Controls
{
    /// <summary>
    /// Custom button with rounded corners, hover/press effects and disabled state.
    /// Validates: Requirements 4.1-4.8
    /// </summary>
    public class ModernButton : Button
    {
        public enum ButtonStyle { Primary, Secondary, Accent }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        public ButtonStyle Style { get; set; } = ButtonStyle.Primary;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [System.ComponentModel.Browsable(false)]
        public bool IsOutline { get; set; } = false;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [System.ComponentModel.Browsable(false)]
        public Color OutlineColor { get; set; } = Color.FromArgb(0x06, 0x85, 0x34);

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [System.ComponentModel.Browsable(false)]
        public Color OutlineHoverFill { get; set; } = Color.FromArgb(0xED, 0xF6, 0xF0);

        private bool _isHovered;
        private bool _isPressed;

        // Base colors set by ThemeManager
        private Color _baseBackColor = Color.FromArgb(0x06, 0x85, 0x34);
        private Color _baseForeColor = Color.White;

        public ModernButton()
        {
            SetStyle(ControlStyles.UserPaint | ControlStyles.AllPaintingInWmPaint | ControlStyles.DoubleBuffer, true);
            FlatStyle = FlatStyle.Flat;
            FlatAppearance.BorderSize = 0;
            Cursor = Cursors.Hand;
        }

        protected override void OnBackColorChanged(EventArgs e)
        {
            base.OnBackColorChanged(e);
            _baseBackColor = BackColor;
        }

        protected override void OnForeColorChanged(EventArgs e)
        {
            base.OnForeColorChanged(e);
            _baseForeColor = ForeColor;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new Rectangle(0, 0, Width - 1, Height - 1);

            if (IsOutline)
            {
                Color borderColor = Enabled ? OutlineColor : Color.FromArgb(0xBE, 0xDF, 0xCA);
                Color fillColor = Enabled ? (_isPressed ? AdjustBrightness(OutlineHoverFill, -0.05f) : (_isHovered ? OutlineHoverFill : Color.White)) : Color.White;
                Color textColor = Enabled ? borderColor : Color.FromArgb(0x7C, 0xBF, 0x94);

                using (var path = GetRoundedPath(rect, 8))
                using (var fillBrush = new SolidBrush(fillColor))
                using (var pen = new Pen(borderColor, 1.5f))
                {
                    e.Graphics.FillPath(fillBrush, path);
                    e.Graphics.DrawPath(pen, path);
                }

                var outlineTextRect = new Rectangle(Padding.Left, 0, Width - Padding.Left - Padding.Right, Height);
                using (var textBrush = new SolidBrush(textColor))
                {
                    var sf = new StringFormat { Alignment = StringAlignment.Center, LineAlignment = StringAlignment.Center };
                    e.Graphics.DrawString(Text, Font, textBrush, outlineTextRect, sf);
                }
                return;
            }

            Color bgColor;
            Color fgColor;

            if (!Enabled)
            {
                bgColor = Color.FromArgb(0xBE, 0xDF, 0xCA);
                fgColor = Color.FromArgb(0x7C, 0xBF, 0x94);
            }
            else if (_isPressed)
            {
                bgColor = AdjustBrightness(_baseBackColor, -0.20f);
                fgColor = _baseForeColor;
            }
            else if (_isHovered)
            {
                bgColor = AdjustBrightness(_baseBackColor, 0.15f);
                fgColor = _baseForeColor;
            }
            else
            {
                bgColor = _baseBackColor;
                fgColor = _baseForeColor;
            }

            using (var path = GetRoundedPath(rect, 8))
            using (var brush = new SolidBrush(bgColor))
            {
                e.Graphics.FillPath(brush, path);
            }

            var textRect = new Rectangle(Padding.Left, 0, Width - Padding.Left - Padding.Right, Height);
            using (var textBrush = new SolidBrush(fgColor))
            {
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Center,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(Text, Font, textBrush, textRect, sf);
            }
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            _isHovered = true;
            Invalidate();
            base.OnMouseEnter(e);
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            _isHovered = false;
            _isPressed = false;
            Invalidate();
            base.OnMouseLeave(e);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            _isPressed = true;
            Invalidate();
            base.OnMouseDown(e);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            _isPressed = false;
            Invalidate();
            base.OnMouseUp(e);
        }

        private static GraphicsPath GetRoundedPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
            path.AddArc(rect.X, rect.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        private static Color AdjustBrightness(Color color, float factor)
        {
            float r = Math.Max(0, Math.Min(255, color.R + 255 * factor));
            float g = Math.Max(0, Math.Min(255, color.G + 255 * factor));
            float b = Math.Max(0, Math.Min(255, color.B + 255 * factor));
            return Color.FromArgb(color.A, (int)r, (int)g, (int)b);
        }
    }
}
