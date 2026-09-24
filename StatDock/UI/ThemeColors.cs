using System;
using System.Drawing;
using System.Windows.Forms;

namespace StatDock
{
    /// <summary>Panel colors that follow the Excel theme (White / Colorful / Dark Gray / Black).</summary>
    public sealed class ThemeColors
    {
        public Color Back, Fore, Dim, ChipBack, ChipHover, ChipBorder, Accent, Warn;
        public bool Dark;

        public static ThemeColors FromBack(Color back)
        {
            var t = new ThemeColors();
            t.Back = back;
            t.Dark = (0.299 * back.R + 0.587 * back.G + 0.114 * back.B) < 128;

            Color toward = t.Dark ? Color.White : Color.Black;
            t.Fore = t.Dark ? Color.FromArgb(242, 242, 242) : Color.FromArgb(38, 38, 38);
            t.Dim = Blend(t.Fore, back, 0.45);
            t.ChipBack = Blend(back, toward, 0.06);
            t.ChipHover = Blend(back, toward, 0.13);
            t.ChipBorder = Blend(back, toward, 0.20);
            t.Accent = t.Dark ? Color.FromArgb(16, 124, 65) : Color.FromArgb(33, 115, 70);   // Excel green
            t.Warn = t.Dark ? Color.FromArgb(255, 170, 120) : Color.FromArgb(180, 70, 20);
            return t;
        }

        public static Color Blend(Color from, Color to, double t)
        {
            return Color.FromArgb(
                (int)(from.R + (to.R - from.R) * t),
                (int)(from.G + (to.G - from.G) * t),
                (int)(from.B + (to.B - from.B) * t));
        }

        /// <summary>Initial estimate from the Office registry, used before a pixel sample is available.</summary>
        public static ThemeColors Guess()
        {
            return FromBack(GuessBackFromRegistry());
        }

        /// <summary>
        /// Read the background color from an Excel status bar pixel. Returns null if the pixel can't be trusted
        /// (that point is covered by another application's window), so the theme doesn't flip incorrectly.
        /// </summary>
        public static ThemeColors TrySample(Point screenPt, IntPtr excelRoot)
        {
            try
            {
                if (!Native.IsInsideRoot(screenPt, excelRoot)) return null;

                using (var bmp = new Bitmap(1, 1))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(screenPt, Point.Empty, new Size(1, 1));
                    return FromBack(bmp.GetPixel(0, 0));
                }
            }
            catch
            {
                return null;
            }
        }

        private static Color GuessBackFromRegistry()
        {
            try
            {
                string ver = "16.0";
                try
                {
                    var v = Globals.Application?.Version;
                    if (!string.IsNullOrEmpty(v)) ver = v;
                }
                catch { }

                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Office\" + ver + @"\Common"))
                {
                    object val = key?.GetValue("UI Theme");
                    if (val is int i)
                    {
                        if (i == 3) return Color.FromArgb(68, 68, 68);    // Dark Gray
                        if (i == 4) return Color.FromArgb(38, 38, 38);    // Black
                        if (i == 6 && IsWindowsDarkMode()) return Color.FromArgb(38, 38, 38);
                    }
                }
            }
            catch { }

            return Color.White;
        }

        private static bool IsWindowsDarkMode()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    object val = key?.GetValue("AppsUseLightTheme");
                    if (val is int i && i == 0) return true;
                }
            }
            catch { }
            return false;
        }
    }
}
