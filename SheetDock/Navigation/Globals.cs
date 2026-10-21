using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetDock
{
    public static class Globals
    {
        public static Excel.Application Application { get; set; }
        public static RecentSheetStack Recent { get; set; } = new RecentSheetStack();

        /// <summary>
        /// True while Excel is in copy / cut mode (marching ants). Reading CutCopyMode does not
        /// cancel it. Anything that Activates / Selects / writes through COM while this is true
        /// would cancel the copy, so the add-in must stay passive in that state.
        /// </summary>
        public static bool CopyModeActive()
        {
            try
            {
                var app = Application;
                return app != null && (int)app.CutCopyMode != 0;
            }
            catch { return false; }
        }

        public static void Log(string text)
        {
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SheetDock.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + text + Environment.NewLine);
            }
            catch { }
        }
    }
}
