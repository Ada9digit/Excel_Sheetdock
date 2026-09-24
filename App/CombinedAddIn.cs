using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetDock
{
    public class CombinedAddIn : IExcelAddIn
    {
        private static readonly Dictionary<int, CustomTaskPane> _ctpPerWindow = new Dictionary<int, CustomTaskPane>();
        private static readonly HashSet<int> _pendingCreate = new HashSet<int>();
        private static readonly Dictionary<int, MsoCTPDockPosition> _lastDock = new Dictionary<int, MsoCTPDockPosition>();

        public void AutoOpen()
        {
            var excelApp = (Excel.Application)ExcelDnaUtil.Application;
            Globals.Application = excelApp;
            StatDock.Globals.Application = excelApp;
            Globals.Recent.Clear();

            var ev = (Excel.AppEvents_Event)excelApp;

            ev.NewWorkbook += App_NewWorkbook;
            ev.WorkbookOpen += App_WorkbookOpen;
            ev.WorkbookActivate += App_WorkbookActivate;
            ev.WindowActivate += App_WindowActivate;
            ev.SheetActivate += App_SheetActivate;
            NavigatorPaneHost.HostResized += OnHostResized;

            ev.SheetSelectionChange += Stat_SheetSelectionChange;
            ev.SheetChange += Stat_SheetChange;
            ev.SheetCalculate += Stat_SheetCalculate;
            ev.SheetActivate += Stat_SheetActivate;
            ev.WorkbookActivate += Stat_WorkbookActivate;
            ev.WindowActivate += Stat_WindowActivate;
            ev.NewWorkbook += Stat_NewWorkbook;
            ev.WorkbookOpen += Stat_WorkbookOpen;

            StatDock.StatHost.Startup();

            EnsureCtpForWindow(excelApp.ActiveWindow);
            HideSheetTabs(excelApp.ActiveWindow);
        }

        public void AutoClose()
        {
            var app = Globals.Application;
            if (app != null)
            {
                var ev = (Excel.AppEvents_Event)app;
                ev.NewWorkbook -= App_NewWorkbook;
                ev.WorkbookOpen -= App_WorkbookOpen;
                ev.WorkbookActivate -= App_WorkbookActivate;
                ev.WindowActivate -= App_WindowActivate;
                ev.SheetActivate -= App_SheetActivate;
                NavigatorPaneHost.HostResized -= OnHostResized;

                ev.SheetSelectionChange -= Stat_SheetSelectionChange;
                ev.SheetChange -= Stat_SheetChange;
                ev.SheetCalculate -= Stat_SheetCalculate;
                ev.SheetActivate -= Stat_SheetActivate;
                ev.WorkbookActivate -= Stat_WorkbookActivate;
                ev.WindowActivate -= Stat_WindowActivate;
                ev.NewWorkbook -= Stat_NewWorkbook;
                ev.WorkbookOpen -= Stat_WorkbookOpen;

                try
                {
                    foreach (Excel.Window wn in app.Windows)
                        try { wn.DisplayWorkbookTabs = true; } catch { }
                }
                catch { }
            }

            StatDock.StatHost.Shutdown();

            foreach (var kv in _ctpPerWindow.ToArray())
                try { kv.Value.Visible = false; kv.Value.Delete(); } catch { }

            _ctpPerWindow.Clear();
            _pendingCreate.Clear();
            _lastDock.Clear();
            NavigatorPane.ShutdownStatics();
        }

        // ---------- SheetDock API ----------
        public static bool IsPaneVisible
        {
            get
            {
                try
                {
                    var wn = Globals.Application?.ActiveWindow;
                    return wn != null && TryGetLiveCtp(wn.Hwnd, out var ctp) && ctp.Visible;
                }
                catch { return false; }
            }
        }

        public static void TogglePane()
        {
            try
            {
                var wn = Globals.Application?.ActiveWindow;
                if (wn == null) return;
                int hwnd = wn.Hwnd;
                if (TryGetLiveCtp(hwnd, out var ctp)) ctp.Visible = !ctp.Visible;
                else CreateCtp(hwnd);
            }
            catch { }
        }

        private static bool TryGetLiveCtp(int hwnd, out CustomTaskPane ctp)
        {
            if (_ctpPerWindow.TryGetValue(hwnd, out ctp) && ctp != null)
            {
                try { bool probe = ctp.Visible; return true; }
                catch
                {
                    _ctpPerWindow.Remove(hwnd);
                    _lastDock.Remove(hwnd);
                }
            }
            ctp = null;
            return false;
        }

        private static void EnsureCtpForWindow(Excel.Window wn)
        {
            if (wn == null) return;
            int hwnd;
            try { hwnd = wn.Hwnd; } catch { return; }
            if (TryGetLiveCtp(hwnd, out _)) return;
            if (!_pendingCreate.Add(hwnd)) return;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { CreateCtp(hwnd); }
                finally { _pendingCreate.Remove(hwnd); }
            });
        }

        private static void CreateCtp(int hwnd)
        {
            try
            {
                var app = Globals.Application;
                if (app == null) return;
                Excel.Window active = null;
                try { active = app.ActiveWindow; } catch { }
                if (active == null || active.Hwnd != hwnd) return;
                if (TryGetLiveCtp(hwnd, out _)) return;

                var ctp = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(NavigatorPaneHost), "SheetDock");
                if (ctp == null) return;
                try { ctp.DockPosition = MsoCTPDockPosition.msoCTPDockPositionLeft; } catch { }
                try { ctp.Width = 220; } catch { }
                ctp.Visible = true;
                _ctpPerWindow[hwnd] = ctp;
                _lastDock[hwnd] = MsoCTPDockPosition.msoCTPDockPositionLeft;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("CreateCtp error: " + ex.Message);
            }
        }

        private static void OnHostResized()
        {
            foreach (var kv in _ctpPerWindow.ToArray())
            {
                try
                {
                    var ctp = kv.Value;
                    var pos = ctp.DockPosition;
                    if (_lastDock.TryGetValue(kv.Key, out var last) && last == pos) continue;
                    _lastDock[kv.Key] = pos;
                    if (pos == MsoCTPDockPosition.msoCTPDockPositionTop ||
                        pos == MsoCTPDockPosition.msoCTPDockPositionBottom)
                        ctp.Height = 52;
                    else if (pos == MsoCTPDockPosition.msoCTPDockPositionLeft ||
                             pos == MsoCTPDockPosition.msoCTPDockPositionRight)
                        ctp.Width = 220;
                }
                catch { }
            }
        }

        private static void HideSheetTabs(Excel.Window wn)
        {
            try { if (wn != null) wn.DisplayWorkbookTabs = false; } catch { }
        }

        private void App_NewWorkbook(Excel.Workbook wb)
        {
            var wn = Globals.Application?.ActiveWindow;
            HideSheetTabs(wn); EnsureCtpForWindow(wn); NavigatorPane.RequestRefreshAll();
        }

        private void App_WorkbookOpen(Excel.Workbook wb)
        {
            var wn = Globals.Application?.ActiveWindow;
            HideSheetTabs(wn); EnsureCtpForWindow(wn); NavigatorPane.RequestRefreshAll();
        }

        private void App_WorkbookActivate(Excel.Workbook wb)
        {
            var wn = Globals.Application?.ActiveWindow;
            HideSheetTabs(wn); EnsureCtpForWindow(wn); NavigatorPane.RequestRefreshAll();
        }

        private void App_SheetActivate(object sh) => NavigatorPane.SyncAllActive();

        private void App_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            try { wn.DisplayWorkbookTabs = false; } catch { }
            EnsureCtpForWindow(wn); NavigatorPane.RequestRefreshAll();
        }

        // ---------- StatDock event router ----------
        private void Stat_SheetSelectionChange(object sh, Excel.Range target) => StatDock.StatHost.RequestRefresh();
        private void Stat_SheetChange(object sh, Excel.Range target) => StatDock.StatHost.RequestRefresh();
        private void Stat_SheetCalculate(object sh) => StatDock.StatHost.RequestRefresh();
        private void Stat_SheetActivate(object sh) => StatDock.StatHost.RequestRefresh();

        private void Stat_WorkbookActivate(Excel.Workbook wb)
        {
            StatDock.StatHost.SuspendVisual(1200);
            StatDock.StatHost.RequestRefresh();
        }

        private void Stat_WindowActivate(Excel.Workbook wb, Excel.Window wn)
        {
            StatDock.StatHost.SuspendVisual(1200);
            StatDock.StatHost.RequestRefresh();
        }

        private void Stat_NewWorkbook(Excel.Workbook wb)
        {
            StatDock.StatHost.SuspendVisual(1500);
            StatDock.StatHost.RequestRefresh();
        }

        private void Stat_WorkbookOpen(Excel.Workbook wb)
        {
            StatDock.StatHost.SuspendVisual(1800);
            StatDock.StatHost.RequestRefresh();
        }
    }
}
