using System;
using System.ComponentModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI.Controls
{
    /// <summary>
    /// A flat, rounded-corner card panel used to give the app a modern,
    /// mobile-like appearance. Supports a subtle hover highlight.
    /// </summary>
    public class RoundedPanel : Panel
    {
        private bool _hovered;
        private Color _borderColor = ThemeManager.ColorBorderLight;
        private Color _hoverBorderColor = ThemeManager.ColorSecondaryAlt;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public int CornerRadius { get; set; } = 16;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public Color? AccentBarColor { get; set; } = null;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public bool HoverEffect { get; set; } = true;

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public Color BorderColor { get => _borderColor; set { _borderColor = value; Invalidate(); } }

        [DesignerSerializationVisibility(DesignerSerializationVisibility.Hidden)]
        [Browsable(false)]
        public Color HoverBorderColor { get => _hoverBorderColor; set { _hoverBorderColor = value; Invalidate(); } }

        public RoundedPanel()
        {
            DoubleBuffered = true;
            BackColor = Color.White;
            SetStyle(ControlStyles.SupportsTransparentBackColor, true);
            MouseEnter += (s, e) => { if (HoverEffect) { _hovered = true; Invalidate(); } };
            MouseLeave += (s, e) => { if (HoverEffect) { _hovered = false; Invalidate(); } };
        }

        private GraphicsPath GetRoundedRect(Rectangle bounds, int radius)
        {
            int d = radius * 2;
            var path = new GraphicsPath();
            if (d <= 0 || bounds.Width < d || bounds.Height < d)
            {
                path.AddRectangle(bounds);
                return path;
            }
            path.AddArc(bounds.X, bounds.Y, d, d, 180, 90);
            path.AddArc(bounds.Right - d, bounds.Y, d, d, 270, 90);
            path.AddArc(bounds.Right - d, bounds.Bottom - d, d, d, 0, 90);
            path.AddArc(bounds.X, bounds.Bottom - d, d, d, 90, 90);
            path.CloseFigure();
            return path;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(0, 0, Width - 1, Height - 1);
            using var path = GetRoundedRect(rect, CornerRadius);

            // Fill parent background outside the rounded shape so corners look transparent
            using (var parentBrush = new SolidBrush(Parent?.BackColor ?? ThemeManager.ColorAppBackground))
                e.Graphics.FillRectangle(parentBrush, ClientRectangle);

            using (var brush = new SolidBrush(BackColor))
                e.Graphics.FillPath(brush, path);

            var borderColor = _hovered && HoverEffect ? _hoverBorderColor : _borderColor;
            using (var pen = new Pen(borderColor, _hovered && HoverEffect ? 2 : 1))
                e.Graphics.DrawPath(pen, path);

            if (AccentBarColor.HasValue)
            {
                using var clip = GetRoundedRect(rect, CornerRadius);
                var oldClip = e.Graphics.Clip;
                e.Graphics.SetClip(clip);
                using (var accentBrush = new SolidBrush(AccentBarColor.Value))
                    e.Graphics.FillRectangle(accentBrush, 0, 0, Width, 4);
                e.Graphics.Clip = oldClip;
            }

            base.OnPaint(e);
        }
    }
}
