using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI.Controls
{
    /// <summary>
    /// Custom GroupBox with Carbone header, white background, light grey border and rounded corners.
    /// Validates: Requirements 6.1-6.4
    /// </summary>
    public class ModernGroupBox : GroupBox
    {
        private static readonly Color HeaderBackground = Color.FromArgb(0x39, 0x39, 0x39); // #393939 Carbone
        private static readonly Color HeaderForeground = Color.White;
        private static readonly Color BorderColor      = Color.FromArgb(0xE0, 0xE0, 0xE0); // #E0E0E0
        private static readonly Font  HeaderFont       = new Font("Segoe UI", 10F, FontStyle.Bold);
        private const int CornerRadius = 4;
        private const int HeaderHeight = 28;
        private const int HeaderPadding = 8;

        public ModernGroupBox()
        {
            BackColor = Color.White;
            ForeColor = Color.FromArgb(0x39, 0x39, 0x39);
            Font = new Font("Segoe UI", 9F);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.Clear(BackColor);

            var bounds = new Rectangle(0, 0, Width - 1, Height - 1);

            // Draw rounded border
            using (var borderPath = GetRoundedPath(bounds, CornerRadius))
            using (var borderPen = new Pen(BorderColor, 1))
            {
                e.Graphics.DrawPath(borderPen, borderPath);
            }

            // Draw header background (top strip)
            var headerRect = new Rectangle(1, 1, Width - 2, HeaderHeight);
            using (var headerPath = GetTopRoundedPath(headerRect, CornerRadius))
            using (var headerBrush = new SolidBrush(HeaderBackground))
            {
                e.Graphics.FillPath(headerBrush, headerPath);
            }

            // Draw header text
            var textRect = new Rectangle(HeaderPadding, 1, Width - HeaderPadding * 2, HeaderHeight);
            using (var textBrush = new SolidBrush(HeaderForeground))
            {
                var sf = new StringFormat
                {
                    Alignment = StringAlignment.Near,
                    LineAlignment = StringAlignment.Center
                };
                e.Graphics.DrawString(Text, HeaderFont, textBrush, textRect, sf);
            }
        }

        protected override Padding DefaultPadding =>
            new Padding(8, HeaderHeight + 4, 8, 8);

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

        private static GraphicsPath GetTopRoundedPath(Rectangle rect, int radius)
        {
            var path = new GraphicsPath();
            int d = radius * 2;
            path.AddArc(rect.X, rect.Y, d, d, 180, 90);
            path.AddArc(rect.Right - d, rect.Y, d, d, 270, 90);
            path.AddLine(rect.Right, rect.Bottom, rect.X, rect.Bottom);
            path.CloseFigure();
            return path;
        }
    }
}
