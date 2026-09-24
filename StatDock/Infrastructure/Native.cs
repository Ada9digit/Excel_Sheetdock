using System;
using System.Drawing;
using System.Runtime.InteropServices;

namespace StatDock
{
    internal static class Native
    {
        public const int GA_ROOT = 2;

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X, Y;
            public POINT(int x, int y) { X = x; Y = y; }
        }

        [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr hWnd);
        [DllImport("user32.dll")] public static extern IntPtr GetAncestor(IntPtr hWnd, int flags);
        [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] public static extern IntPtr WindowFromPoint(POINT pt);

        [DllImport("user32.dll")] private static extern bool GetCursorPos(out POINT pt);
        [DllImport("user32.dll")] private static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern short GetAsyncKeyState(int vKey);
        [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr hWnd, out RECT rect);
        [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr hWnd, ref POINT pt);
        [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr hWnd);   // Win10 1607+
        [DllImport("dwmapi.dll")] private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int value, int size);

        /// <summary>True while the left / right mouse button is held down (e.g. while drag-selecting cells).</summary>
        public static bool MouseButtonDown()
        {
            return (GetAsyncKeyState(0x01) & 0x8000) != 0 || (GetAsyncKeyState(0x02) & 0x8000) != 0;
        }

        /// <summary>True if the mouse cursor is inside that window.</summary>
        public static bool CursorInWindow(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return false;

            POINT p;
            RECT r;
            if (!GetCursorPos(out p) || !GetWindowRect(hwnd, out r)) return false;

            return p.X >= r.Left && p.X < r.Right && p.Y >= r.Top && p.Y < r.Bottom;
        }

        /// <summary>Window DPI (96 = 100%). Falls back to 96 on older Windows.</summary>
        public static int DpiOf(IntPtr hwnd)
        {
            try
            {
                uint d = GetDpiForWindow(hwnd);
                if (d > 0) return (int)d;
            }
            catch { }
            return 96;
        }

        /// <summary>A window's client area in screen coordinates.</summary>
        public static bool GetClientScreenRect(IntPtr hwnd, out Rectangle rect)
        {
            rect = Rectangle.Empty;

            RECT rc;
            if (!GetClientRect(hwnd, out rc)) return false;

            var pt = new POINT(0, 0);
            if (!ClientToScreen(hwnd, ref pt)) return false;

            rect = new Rectangle(pt.X, pt.Y, rc.Right - rc.Left, rc.Bottom - rc.Top);
            return rect.Width > 0 && rect.Height > 0;
        }

        /// <summary>True if the pixel at that screen point belongs to an Excel window (not another application drawn on top).</summary>
        public static bool IsInsideRoot(Point screenPt, IntPtr root)
        {
            IntPtr h = WindowFromPoint(new POINT(screenPt.X, screenPt.Y));
            return h != IntPtr.Zero && GetAncestor(h, GA_ROOT) == root;
        }

        /// <summary>Windows 11: disable rounded corners and the thin border so the strip blends into the status bar.</summary>
        public static void SquareOff(IntPtr hwnd)
        {
            try
            {
                int doNotRound = 1;                 // DWMWCP_DONOTROUND
                DwmSetWindowAttribute(hwnd, 33, ref doNotRound, sizeof(int));   // DWMWA_WINDOW_CORNER_PREFERENCE

                int none = unchecked((int)0xFFFFFFFE);                            // DWMWA_COLOR_NONE
                DwmSetWindowAttribute(hwnd, 34, ref none, sizeof(int));         // DWMWA_BORDER_COLOR
            }
            catch { }
        }
    }
}
