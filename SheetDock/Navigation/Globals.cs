using Excel = Microsoft.Office.Interop.Excel;

namespace SheetDock
{
    public static class Globals
    {
        public static Excel.Application Application { get; set; }
        public static RecentSheetStack Recent { get; set; } = new RecentSheetStack();
    }
}
