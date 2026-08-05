using System;
using System.Drawing;
using System.Windows.Forms;

namespace AuserExcelTransformer.UI
{
    public static class ThemeManager
    {
        // Palette (Auser modern green palette)
        public static readonly Color ColorBackground   = Color.White;
        public static readonly Color ColorAppBackground= Color.FromArgb(0xED, 0xF6, 0xF0); // #edf6f0 Soft mint background
        public static readonly Color ColorPrimary      = Color.FromArgb(0x06, 0x85, 0x34); // #068534 Deep green (was Carbone)
        public static readonly Color ColorAccent       = Color.FromArgb(0x06, 0x85, 0x34); // #068534 Deep green
        public static readonly Color ColorAccentHover  = Color.FromArgb(0x27, 0x95, 0x4F); // #27954f
        public static readonly Color ColorSecondary    = Color.FromArgb(0x3B, 0x9F, 0x5F); // #3b9f5f Mid green
        public static readonly Color ColorSecondaryAlt = Color.FromArgb(0x56, 0xAC, 0x75); // #56ac75
        public static readonly Color ColorTertiary     = Color.FromArgb(0x7C, 0xBF, 0x94); // #7cbf94
        public static readonly Color ColorSoft         = Color.FromArgb(0x9E, 0xCF, 0xB0); // #9ecfb0
        public static readonly Color ColorDisabled     = Color.FromArgb(0xBE, 0xDF, 0xCA); // #bedfca
        public static readonly Color ColorDisabledText = Color.FromArgb(0x7C, 0xBF, 0x94); // #7cbf94
        public static readonly Color ColorBorderLight  = Color.FromArgb(0xBE, 0xDF, 0xCA); // #bedfca
        public static readonly Color ColorRowAlt       = Color.FromArgb(0xED, 0xF6, 0xF0); // #edf6f0
        public static readonly Color ColorError        = Color.FromArgb(0xD3, 0x2F, 0x2F); // #D32F2F (kept for error states)
        public static readonly Color ColorHeaderText   = Color.White;

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
            btn.BackColor = ColorSecondaryAlt;
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
            btn.BackColor = ColorSoft;
            btn.ForeColor = ColorPrimary;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        // Outline variant: transparent/white fill with a colored border — used for
        // secondary actions (Importa, Cancella) per the inspiration design.
        public static void ApplyOutlinePrimary(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.IsOutline = true;
            btn.OutlineColor = ColorPrimary;
            btn.OutlineHoverFill = ColorAppBackground;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        public static void ApplyOutlineNeutral(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.IsOutline = true;
            btn.OutlineColor = Color.FromArgb(0x6E, 0x7A, 0x6C);
            btn.OutlineHoverFill = ColorAppBackground;
            btn.Font = FontNormal;
            btn.MinimumSize = new System.Drawing.Size(0, 40);
            btn.Padding = new Padding(20, 0, 20, 0);
            btn.FlatStyle = FlatStyle.Flat;
            btn.FlatAppearance.BorderSize = 0;
        }

        public static void ApplyOutlineError(Controls.ModernButton btn)
        {
            if (btn == null) return;
            btn.IsOutline = true;
            btn.OutlineColor = ColorError;
            btn.OutlineHoverFill = Color.FromArgb(0xFF, 0xDA, 0xD6);
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
