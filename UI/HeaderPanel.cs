using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI
{
    /// <summary>
    /// Header panel with fixed 80px height, Carbone background, logo and title.
    /// Validates: Requirements 2.1-2.5, 10.3-10.4
    /// </summary>
    public class HeaderPanel : Panel
    {
        private Image? _logo;
        private readonly Label _titleLabel;

        public HeaderPanel()
        {
            Height = 80;
            BackColor = Color.FromArgb(0x39, 0x39, 0x39); // Carbone
            Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            Dock = DockStyle.None;

            _titleLabel = new Label
            {
                Text = "Auser Gestione Trasporti",
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 18F, FontStyle.Bold),
                AutoSize = false,
                TextAlign = ContentAlignment.MiddleLeft,
                BackColor = Color.Transparent
            };
            Controls.Add(_titleLabel);

            LoadLogo();
            PositionControls();
        }

        private void LoadLogo()
        {
            try
            {
                var logoPath = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "logo", "Auser_logo.png");
                if (System.IO.File.Exists(logoPath))
                {
                    var original = Image.FromFile(logoPath);
                    // Scale proportionally to max height 60px
                    int maxH = 60;
                    int newH = Math.Min(original.Height, maxH);
                    int newW = (int)((double)original.Width / original.Height * newH);
                    var scaled = new Bitmap(newW, newH);
                    using (var g = Graphics.FromImage(scaled))
                    {
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.DrawImage(original, 0, 0, newW, newH);
                    }
                    _logo = scaled;
                    original.Dispose();
                }
            }
            catch
            {
                // Graceful fallback: no logo, title only
                _logo = null;
            }
        }

        private void PositionControls()
        {
            if (_titleLabel == null) return;

            int logoWidth = _logo != null ? _logo.Width + 20 : 0;
            int logoLeft = 20;
            int titleLeft = logoLeft + logoWidth + (_logo != null ? 10 : 0);
            int titleWidth = Math.Max(10, Width - titleLeft - 10);

            _titleLabel.Location = new Point(titleLeft, 0);
            _titleLabel.Size = new Size(titleWidth, Height);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            PositionControls();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (_logo != null)
            {
                int logoY = (Height - _logo.Height) / 2;
                e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
                e.Graphics.DrawImage(_logo, 20, logoY, _logo.Width, _logo.Height);
            }
        }
    }
}
