using System;
using System.Drawing;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI
{
    public static class ThemeManager
    {
        // Palette
        public static readonly Color ColorBackground   = Color.White;
        public static readonly Color ColorPrimary      = Color.FromArgb(0x39, 0x39, 0x39); // #393939 Carbone
        public static readonly Color ColorAccent       = Color.FromArgb(0x00, 0x92, 0x46); // #009246 Verde
        public static readonly Color ColorSecondary    = Color.FromArgb(0xFA, 0xB9, 0x00); // #FAB900 Ambra
        public static readonly Color ColorDisabled     = Color.FromArgb(0xCC, 0xCC, 0xCC);
        public static readonly Color ColorDisabledText = Color.FromArgb(0x88, 0x88, 0x88);
        public static readonly Color ColorBorderLight  = Color.FromArgb(0xE0, 0xE0, 0xE0);
        public static readonly Color ColorRowAlt       = Color.FromArgb(0xF5, 0xF5, 0xF5);
        public static readonly Color ColorError        = Color.FromArgb(0xD3, 0x2F, 0x2F); // #D32F2F

        // Font
        public static readonly Font FontTitle        = new Font("Segoe UI", 24F, FontStyle.Bold);
        public static readonly Font FontSubtitle     = new Font("Segoe UI", 12F, FontStyle.Bold);
        public static readonly Font FontNormal       = new Font("Segoe UI", 10F);
        public static readonly Font FontSmall        = new Font("Segoe UI", 9F);
        public static readonly Font FontSectionLabel = new Font("Segoe UI", 14F, FontStyle.Bold);
        public static readonly Font FontGroupHeader  = new Font("Segoe UI", 10F, FontStyle.Bold);

        public static void ApplyPrimary(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.Style = Controls.ModernButton.ButtonStyle.Primary;
            btn.BackColor = ColorAccent;
            btn.ForeColor = ColorBackground;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        public static void ApplySecondary(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.Style = Controls.ModernButton.ButtonStyle.Secondary;
            btn.BackColor = ColorPrimary;
            btn.ForeColor = ColorBackground;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        public static void ApplyAccent(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.Style = Controls.ModernButton.ButtonStyle.Accent;
            btn.BackColor = ColorSecondary;
            btn.ForeColor = ColorPrimary;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        public static void ApplyStyle(Label lbl)
        {
            if (lbl == null) return;
            lbl.Font = FontSmall;
            lbl.ForeColor = ColorPrimary;
            lbl.BackColor = Color.Transparent;
        }

        public static void ApplyStyle(ListView lv)
        {
            if (lv == null) return;
            lv.Font = FontSmall;
            lv.BackColor = ColorBackground;
            lv.ForeColor = ColorPrimary;
            // Column header styling via OwnerDraw
            lv.OwnerDraw = true;
            lv.DrawColumnHeader += (s, e) =>
            {
                using var brush = new System.Drawing.SolidBrush(ColorPrimary);
                e.Graphics.FillRectangle(brush, e.Bounds);
                using var textBrush = new System.Drawing.SolidBrush(ColorBackground);
                var fmt = new System.Drawing.StringFormat { Alignment = System.Drawing.StringAlignment.Near, LineAlignment = System.Drawing.StringAlignment.Center };
                e.Graphics.DrawString(e.Header.Text, FontSmall, textBrush, e.Bounds, fmt);
            };
            lv.DrawItem += (s, e) =>
            {
                e.DrawDefault = true;
            };
            lv.DrawSubItem += (s, e) =>
            {
                var bg = e.ItemIndex % 2 == 0 ? ColorBackground : ColorRowAlt;
                using var brush = new System.Drawing.SolidBrush(bg);
                e.Graphics.FillRectangle(brush, e.Bounds);
                using var textBrush = new System.Drawing.SolidBrush(ColorPrimary);
                var fmt = new System.Drawing.StringFormat { Alignment = System.Drawing.StringAlignment.Near, LineAlignment = System.Drawing.StringAlignment.Center };
                var textBounds = new System.Drawing.Rectangle(e.Bounds.X + 2, e.Bounds.Y, e.Bounds.Width - 2, e.Bounds.Height);
                e.Graphics.DrawString(e.SubItem.Text, FontSmall, textBrush, textBounds, fmt);
            };
        }

        public static void ApplyStyle(ComboBox cmb)
        {
            if (cmb == null) return;
            cmb.Font = FontSmall;
            cmb.BackColor = ColorBackground;
            cmb.ForeColor = ColorPrimary;
            cmb.FlatStyle = FlatStyle.Flat;
        }

        public static void ApplyStyle(ProgressBar pb)
        {
            if (pb == null) return;
            pb.BackColor = ColorBorderLight;
            // ProgressBar ForeColor is not directly supported on Windows; use SetWindowTheme workaround
            try
            {
                NativeMethods.SetWindowTheme(pb.Handle, "", "");
                pb.ForeColor = ColorAccent;
            }
            catch { /* ignore if not available */ }
        }
    }

    internal static class NativeMethods
    {
        [System.Runtime.InteropServices.DllImport("uxtheme.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        internal static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);
    }
}
