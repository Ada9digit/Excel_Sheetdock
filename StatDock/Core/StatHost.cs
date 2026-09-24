using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatDock
{
    /// <summary>
    /// Single coordinator: one StatBar per Excel main window (Excel 2013+ opens each workbook
    /// in its own window), one selection read for the active bar, and three lightweight timers:
    ///  - debounce : merges consecutive events into a single read
    ///  - track    : attaches the strip to the status bar (pure P/Invoke, no COM)
    ///  - poll     : detects filter / row-hide changes that don't raise Excel events
    /// </summary>
    internal static class StatHost
    {
        private static readonly Dictionary<IntPtr, StatBar> Bars = new Dictionary<IntPtr, StatBar>();

        private static Timer _debounce;
        private static Timer _track;
        private static Timer _poll;

        private static bool _enabled = true;      // toggled from the ribbon
        private static bool _statusBarOn = true;  // Application.DisplayStatusBar
        private static int _lastCaptureMs;
        private static DateTime _nextPollAt = DateTime.MinValue;
        private static string _lastCaptureSig = "";
        private static string _lastPollSig = "";
        private static IntPtr _gridHwnd = IntPtr.Zero;   // active workbook window (sheet area)
        private static bool _pendingAfterMouse;   // refresh deferred while a mouse button is held down
        private static bool _updateRunning;
        private static DateTime _suspendVisualUntil = DateTime.MinValue;
        private static DateTime _eventMuteUntil = DateTime.MinValue;
        private static DateTime _busyRetryAt = DateTime.MinValue;
        private static DateTime _filterSettleUntil = DateTime.MinValue;
        // Hard gate: prevents Excel events emitted by Capture/ApplySnapshot/native Paste
        // from immediately re-entering another Capture cycle.
        private static DateTime _nextUpdateAt = DateTime.MinValue;
        private static bool _externalEditRefreshPending;
        private static bool _refreshPending;

        /// <summary>Called from AutoOpen (Excel's main thread).</summary>
        public static void Startup()
        {
            if (_debounce != null) return;

            // Don't show the overlay right after Excel finishes loading the add-in / workbook.
            _suspendVisualUntil = DateTime.UtcNow.AddMilliseconds(1800);

            _debounce = new Timer { Interval = 100 };
            _debounce.Tick += (s, e) =>
            {
                _debounce.Stop();

                // External Excel operation: wait until the mute window has really expired.
                if (_externalEditRefreshPending && DateTime.UtcNow < _eventMuteUntil)
                {
                    var remaining = (int)Math.Max(50, (_eventMuteUntil - DateTime.UtcNow).TotalMilliseconds);
                    _debounce.Interval = Math.Min(2000, remaining);
                    _debounce.Start();
                    return;
                }

                if (DateTime.UtcNow < _eventMuteUntil ||
                    DateTime.UtcNow < _filterSettleUntil ||
                    DateTime.UtcNow < _nextUpdateAt ||
                    _updateRunning)
                {
                    _debounce.Interval = 100;
                    _debounce.Start();
                    return;
                }

                _debounce.Interval = 100;
                _externalEditRefreshPending = false;
                _refreshPending = false;
                Safe(Update);
            };

            _track = new Timer { Interval = 50 };
            _track.Tick += (s, e) => Safe(Track);
            _track.Start();

            _poll = new Timer { Interval = 1000 };
            _poll.Tick += (s, e) => Safe(Poll);
            _poll.Start();

            RequestRefresh();
        }

        public static void Shutdown()
        {
            try
            {
                if (_debounce != null) { _debounce.Stop(); _debounce.Dispose(); }
                if (_track != null) { _track.Stop(); _track.Dispose(); }
                if (_poll != null) { _poll.Stop(); _poll.Dispose(); }
            }
            catch { }

            _debounce = null;
            _track = null;
            _poll = null;

            foreach (var bar in Bars.Values)
            {
                try { bar.Dispose(); }
                catch { }
            }
            Bars.Clear();
        }

        /// <summary>Request a refresh (coalesced: consecutive events become a single read).</summary>
        /// <summary>
        /// Mute refresh events while StatDock causes an Excel UI operation (for example native Paste).
        /// Excel emits several selection/change/calculate events synchronously; those events must not
        /// start nested StatEngine.Capture calls.
        /// </summary>
        public static void BeginExternalEdit()
        {
            _externalEditRefreshPending = true;
            _eventMuteUntil = DateTime.UtcNow.AddSeconds(2);
            _filterSettleUntil = DateTime.UtcNow.AddMilliseconds(900);
            _nextUpdateAt = _filterSettleUntil;

            if (_debounce != null)
                _debounce.Stop();
        }

        /// <summary>
        /// End an Excel UI operation and schedule exactly one StatDock refresh after Excel settles.
        /// </summary>
        public static void EndExternalEdit(int settleMilliseconds)
        {
            if (settleMilliseconds < 500) settleMilliseconds = 500;

            _externalEditRefreshPending = true;
            _eventMuteUntil = DateTime.UtcNow.AddMilliseconds(settleMilliseconds);
            _filterSettleUntil = DateTime.UtcNow.AddMilliseconds(settleMilliseconds);
            _nextUpdateAt = _filterSettleUntil;

            if (_debounce == null) return;

            _debounce.Stop();
            _debounce.Interval = settleMilliseconds;
            _debounce.Start();
        }

        public static void RequestRefresh()
        {
            if (_debounce == null) return;

            if (Native.MouseButtonDown())
            {
                _pendingAfterMouse = true;
                _refreshPending = true;
                return;
            }

            var now = DateTime.UtcNow;

            if (now < _eventMuteUntil ||
                now < _filterSettleUntil ||
                now < _nextUpdateAt ||
                _updateRunning)
            {
                // Do not queue another refresh. Polling will detect a real post-operation
                // change after the gate expires. This is the important anti-loop guard.
                return;
            }

            _refreshPending = true;
            _debounce.Stop();
            _debounce.Interval = 100;
            _debounce.Start();
        }

        public static void SuspendVisual(int milliseconds)
        {
            if (milliseconds < 0) milliseconds = 0;
            var until = DateTime.UtcNow.AddMilliseconds(milliseconds);
            if (until > _suspendVisualUntil) _suspendVisualUntil = until;

            foreach (var bar in Bars.Values)
            {
                try { bar.SetSuspended(true); }
                catch { }
            }
        }

        /// <summary>Ribbon button: show / hide the strip.</summary>
        public static void Toggle()
        {
            _enabled = !_enabled;
            Safe(Update);
        }

        // ------------------------------------------------------------

        private static void Safe(Action a)
        {
            try { a(); }
            catch (Exception ex)
            {
                Debug.WriteLine("StatHost: " + ex.Message);
            }
        }

        /// <summary>The currently active Excel main window (XLMAIN). Each workbook window has its own XLMAIN.</summary>
        private static IntPtr ActiveRoot(Excel.Application app)
        {
            try
            {
                Excel.Window wn = app.ActiveWindow;
                if (wn != null)
                {
                    _gridHwnd = new IntPtr(wn.Hwnd);
                    IntPtr root = Native.GetAncestor(_gridHwnd, Native.GA_ROOT);
                    StatEngine.Release(wn);
                    if (root != IntPtr.Zero) return root;
                }
            }
            catch { }

            try { return new IntPtr(app.Hwnd); }
            catch { return IntPtr.Zero; }
        }

        /// <summary>Create the bar if needed, attach it to the status bar, then read the selection for the active bar.</summary>
        private static void Update()
        {
            if (_updateRunning) return;

            if (DateTime.UtcNow < _suspendVisualUntil)
            {
                Track();
                return;
            }

            if (Native.MouseButtonDown())
            {
                _pendingAfterMouse = true;
                return;
            }

            _updateRunning = true;
            try
            {
                var app = Globals.Application;
                if (app == null) return;

                // Don't enter SpecialCells/WorksheetFunction while Excel is still calculating
                // or not yet Ready. One scheduled retry is cheaper than dozens of
                // Calculate/Change events triggering each other.
                if (IsExcelBusy(app))
                {
                    _busyRetryAt = DateTime.UtcNow.AddMilliseconds(500);
                    return;
                }

                try { _statusBarOn = app.DisplayStatusBar; }
                catch { }

                IntPtr root = ActiveRoot(app);
                if (_enabled && root != IntPtr.Zero && !Bars.ContainsKey(root))
                    Bars[root] = new StatBar(root);

                Track();

                StatBar bar = null;
                if (!_enabled || !Bars.TryGetValue(root, out bar) || !bar.Visible || bar.DragActive) return;

                var sw = Stopwatch.StartNew();
                SelectionSnapshot snap = null;
                try { snap = StatEngine.Capture(app, false); }
                catch { }
                _lastCaptureMs = (int)sw.ElapsedMilliseconds;
                if (snap != null)
                {
                    _lastCaptureSig = snap.Signature;
                    _lastPollSig = snap.Signature;
                }

                bar.ApplyTheme();
                bar.ApplySnapshot(snap);
            }
            finally
            {
                _updateRunning = false;

                // Capture/ApplySnapshot can raise Excel events. Give a short settle window
                // so those events don't form an Update -> Poll -> Update chain.
                _eventMuteUntil = DateTime.UtcNow.AddMilliseconds(450);
                _filterSettleUntil = DateTime.UtcNow.AddMilliseconds(250);
                _nextUpdateAt = DateTime.UtcNow.AddMilliseconds(750);
                _refreshPending = false;
                _nextPollAt = _nextUpdateAt;
            }
        }

        /// <summary>Pure P/Invoke (no COM), so it's safe to call frequently.</summary>
        private static void Track()
        {
            bool suspended = DateTime.UtcNow < _suspendVisualUntil;

            if (_pendingAfterMouse && !Native.MouseButtonDown() && !suspended)
            {
                _pendingAfterMouse = false;
                RequestRefresh();
            }

            if (Bars.Count == 0) return;

            bool show = _enabled && _statusBarOn && !suspended;
            List<IntPtr> dead = null;

            foreach (var kv in Bars)
            {
                try { kv.Value.SetSuspended(suspended); }
                catch { }

                if (!kv.Value.Sync(show))
                {
                    if (dead == null) dead = new List<IntPtr>();
                    dead.Add(kv.Key);
                }
            }

            if (dead == null) return;

            foreach (var key in dead)
            {
                try { Bars[key].Dispose(); }
                catch { }
                Bars.Remove(key);
            }
        }

        /// <summary>
        /// Filtering / hiding rows doesn't raise Excel events, so it's checked periodically. Kept cheap:
        ///  - only while the Excel window that owns the bar is in the foreground (not a dialog / another app)
        ///  - the interval widens automatically if the last read was slow (very large selection)
        /// </summary>
        private static void Poll()
        {
            if (!_enabled || Bars.Count == 0) return;
            if (_updateRunning) return;
            if (DateTime.UtcNow < _nextPollAt) return;
            if (DateTime.UtcNow < _busyRetryAt)
            {
                _nextPollAt = _busyRetryAt;
                return;
            }

            // Don't inspect COM during mouse/filter operations. Excel can send many
            // events within one operation, and every inspection here would only prolong the loop.
            if (Native.MouseButtonDown())
            {
                _pendingAfterMouse = true;
                _nextPollAt = DateTime.UtcNow.AddMilliseconds(250);
                return;
            }

            if (DateTime.UtcNow < _suspendVisualUntil)
            {
                _nextPollAt = _suspendVisualUntil;
                return;
            }

            IntPtr fgRoot = Native.GetAncestor(Native.GetForegroundWindow(), Native.GA_ROOT);
            if (!Bars.ContainsKey(fgRoot))
            {
                _nextPollAt = DateTime.UtcNow.AddMilliseconds(1000);
                return;
            }

            // Periodic COM calls make Excel's cursor flicker to "loading" when the cursor is over the
            // ribbon / status bar. So polling only runs while the cursor is in the sheet area.
            if (!Native.CursorInWindow(_gridHwnd))
            {
                _nextPollAt = DateTime.UtcNow.AddMilliseconds(1000);
                return;
            }

            // Cheap check first (address + height + width). A full read only happens if something changed,
            // e.g. a filter was applied or rows were hidden.
            var app = Globals.Application;
            if (app == null) return;
            if (IsExcelBusy(app))
            {
                _busyRetryAt = DateTime.UtcNow.AddMilliseconds(500);
                _nextPollAt = _busyRetryAt;
                return;
            }

            if (DateTime.UtcNow < _eventMuteUntil ||
                DateTime.UtcNow < _filterSettleUntil ||
                DateTime.UtcNow < _nextUpdateAt)
            {
                _nextPollAt = _nextUpdateAt;
                return;
            }

            string sig = StatEngine.Signature(app);
            if (sig == null)
            {
                _nextPollAt = DateTime.UtcNow.AddMilliseconds(1000);
                return;
            }

            // Poll keeps its own baseline. Capture() doesn't modify this baseline,
            // so internal changes during Update don't become a false loop.
            if (sig == _lastPollSig)
            {
                _nextPollAt = DateTime.UtcNow.AddMilliseconds(1000);
                return;
            }

            _lastPollSig = sig;
            RequestRefresh();

            _nextPollAt = DateTime.UtcNow.AddMilliseconds(
                Math.Max(700, Math.Min(2500, _lastCaptureMs * 20)));
        }

        private static bool IsExcelBusy(Excel.Application app)
        {
            try
            {
                if (!app.Ready) return true;
            }
            catch { return true; }

            try
            {
                return app.CalculationState != Excel.XlCalculationState.xlDone;
            }
            catch { return false; }
        }
    }
}
