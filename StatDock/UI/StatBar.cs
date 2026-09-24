using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatDock
{
    /// <summary>
    /// A thin borderless strip that overlays Excel's status bar (not a task pane).
    /// It is a popup window owned by Excel's main window: always on top of Excel, never
    /// takes focus, and its position is followed by StatHost.Track() when Excel is moved / resized.
    /// It only covers the statistics area on the right side of the status bar (where the native Average / Count / Sum live);
    /// left-side text such as "Ready" and "Sheet 1 of 33" still belongs to Excel. Contains statistic chips
    /// that can be dragged onto a cell as a formula.
    /// </summary>
    internal sealed class StatBar : Form
    {
        // ---------------- Tuning ----------------

        /// <summary>Native status bar height at 100% DPI. Adjust if the strip looks taller / shorter.</summary>
        private const int StatusBarHeightPx = 22;

        /// <summary>
        /// Parts of the native status bar that must not be covered (pixels @100% DPI).
        /// Left 270 = "Ready", "12 of 15 records found", "Sheet 1 of 33". Increase if your left-side text is longer.
        /// Right 300 = view buttons + zoom slider.
        /// </summary>
        private static readonly int KeepNativeLeftPx = 270;
        private static readonly int KeepNativeRightPx = 300;

        /// <summary>Chips hidden first when there isn't enough width (Sum is never hidden).</summary>
        private static readonly StatKind[] DropOrder =
        {
            StatKind.NumCount, StatKind.Count, StatKind.Min, StatKind.Max, StatKind.Average
        };

        private enum Tone { Dim, Accent, Warn }

        private sealed class HwndOwner : IWin32Window
        {
            public HwndOwner(IntPtr handle) { Handle = handle; }
            public IntPtr Handle { get; private set; }
        }

        // ---------------- State ----------------

        private readonly IntPtr _root;   // Excel main window (XLMAIN) that owns this strip
        private readonly Font _font = new Font("Segoe UI", 9f);
        private readonly Dictionary<StatKind, StatChip> _chips = new Dictionary<StatKind, StatChip>();
        private readonly Dictionary<StatKind, string> _tips = new Dictionary<StatKind, string>();
        private readonly ToolTip _tip = new ToolTip();

        private ThemeColors _theme;
        private float _scale = 1f;
        private bool _shownOnce;
        private bool _wantShown;
        private bool _suspended;
        private Rectangle _lastRect = Rectangle.Empty;

        // A short message (drop result / error) replaces the chips for a few seconds.
        private string _flashText = "";
        private Tone _flashTone = Tone.Accent;
        private bool _flashActive;
        private DateTime _flashUntil = DateTime.MinValue;
        private DateTime _nextThemeCheck = DateTime.MinValue;

        // drag
        private SelectionSnapshot _dragSnap;
        private StatKind _dragKind;
        private DragGhost _ghost;
        private string _dragFormulaLocal;
        private int _lastLookup;

        public bool DragActive { get; private set; }

        public StatBar(IntPtr root)
        {
            _root = root;
            _theme = ThemeColors.Guess();

            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
            DoubleBuffered = true;
            Font = _font;
            BackColor = _theme.Back;

            foreach (var kind in StatInfo.Order)
            {
                var chip = new StatChip(kind);
                chip.Theme = _theme;
                chip.DragStarting = OnChipDragStarting;
                chip.DragMoved = OnChipDragMoved;
                chip.DragEnded = OnChipDragEnded;

                _chips[kind] = chip;
                Controls.Add(chip);
            }

            _tip.InitialDelay = 400;
            _tip.ShowAlways = true;   // this form is never active; without this the tooltip won't show
            Size = new Size(600, S(StatusBarHeightPx));
        }

        private int S(int px) { return (int)Math.Round(px * _scale); }

        // ---------------- Window: no focus, on top of Excel ----------------

        protected override bool ShowWithoutActivation { get { return true; } }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x08000000;   // WS_EX_NOACTIVATE
                cp.ExStyle |= 0x00000080;   // WS_EX_TOOLWINDOW (hidden from Alt-Tab)
                return cp;
            }
        }

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            Native.SquareOff(Handle);
        }

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == 0x0021)   // WM_MOUSEACTIVATE -> MA_NOACTIVATE: clicks don't steal focus from Excel
            {
                m.Result = (IntPtr)3;
                return;
            }
            base.WndProc(ref m);
        }

        // ---------------- Position ----------------

        public void SetSuspended(bool suspended)
        {
            _suspended = suspended;
            if (suspended)
            {
                _lastRect = Rectangle.Empty;
                SetShown(false);
            }
        }

        /// <summary>
        /// Align the strip with the status bar of its owning Excel window. Called periodically by StatHost.
        /// Returns false if that Excel window no longer exists (the strip must be disposed).
        /// </summary>
        public bool Sync(bool enabled)
        {
            if (IsDisposed) return false;
            if (!Native.IsWindow(_root)) return false;

            Rectangle client = Rectangle.Empty;
            bool want = enabled &&
                        !_suspended &&
                        Native.IsWindowVisible(_root) &&
                        !Native.IsIconic(_root) &&
                        Native.GetClientScreenRect(_root, out client);

            if (!want)
            {
                SetShown(false);
                return true;
            }

            float scale = Native.DpiOf(_root) / 96f;
            if (Math.Abs(scale - _scale) > 0.01f)
            {
                _scale = scale;
                foreach (var chip in _chips.Values) chip.SetScale(scale);
                LayoutChips();
            }

            if (_flashActive && DateTime.UtcNow >= _flashUntil)
            {
                _flashActive = false;
                LayoutChips();
                Invalidate();
            }

            // Width follows the chip content, anchored to the left of the preserved area (view buttons + zoom).
            int h = S(StatusBarHeightPx);
            int rightEdge = client.Right - S(KeepNativeRightPx);
            int maxWidth = client.Width - S(KeepNativeLeftPx) - S(KeepNativeRightPx);
            int width = Math.Min(ContentWidth(), maxWidth);

            if (width < S(120))
            {
                SetShown(false);
                return true;
            }

            var r = new Rectangle(rightEdge - width, client.Bottom - h, width, h);

            // Only set when the target changes. If Windows rounds / clamps the size, comparing
            // against Bounds would keep differing and trigger SetWindowPos repeatedly every 50 ms.
            if (r != _lastRect)
            {
                _lastRect = r;
                Bounds = r;
            }

            SetShown(true);
            return true;
        }

        private void SetShown(bool want)
        {
            if (IsDisposed || want == _wantShown) return;   // only act when the state changes
            _wantShown = want;

            try
            {
                if (want)
                {
                    if (!_shownOnce)
                    {
                        // Show(owner) makes Excel's main window the owner: the strip minimizes / restores together with Excel.
                        Show(new HwndOwner(_root));
                        _shownOnce = true;
                    }
                    else
                    {
                        Show();
                    }
                }
                else if (Visible)
                {
                    Hide();
                }
            }
            catch
            {
                _wantShown = false;
            }
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            LayoutChips();
            Invalidate();
        }

        // ---------------- Chip layout ----------------

        private int TotalWidth(List<StatKind> kinds, int gap)
        {
            int total = 0;
            foreach (var k in kinds) total += _chips[k].PreferredWidth;
            if (kinds.Count > 1) total += gap * (kinds.Count - 1);
            return total;
        }

        /// <summary>Strip width required to fit all chips.</summary>
        private int ContentWidth()
        {
            return TotalWidth(new List<StatKind>(StatInfo.Order), S(4)) + S(6) * 2;
        }

        /// <summary>Chips are laid out left to right across the strip. If space is tight, low-priority chips are hidden.</summary>
        private void LayoutChips()
        {
            if (_chips.Count == 0) return;

            if (DateTime.UtcNow < _flashUntil)
            {
                foreach (var c in _chips.Values)
                    if (c.Visible) c.Visible = false;
                return;
            }

            int pad = S(6);
            int gap = S(4);
            int chipH = Math.Max(12, ClientSize.Height - S(4));
            int y = (ClientSize.Height - chipH) / 2;
            int avail = ClientSize.Width - pad * 2;

            var shown = new List<StatKind>(StatInfo.Order);
            foreach (var drop in DropOrder)
            {
                if (TotalWidth(shown, gap) <= avail) break;
                shown.Remove(drop);
            }

            int x = pad;
            foreach (var kind in StatInfo.Order)
            {
                StatChip chip = _chips[kind];

                if (!shown.Contains(kind))
                {
                    if (chip.Visible) chip.Visible = false;
                    continue;
                }

                int w = chip.PreferredWidth;
                var b = new Rectangle(x, y, w, chipH);
                if (chip.Bounds != b) chip.Bounds = b;
                if (!chip.Visible) chip.Visible = true;

                x += w + gap;
            }
        }

        // ---------------- Paint (selection info) ----------------

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(_theme.Back);
            if (DateTime.UtcNow >= _flashUntil || string.IsNullOrEmpty(_flashText)) return;

            Color color = _flashTone == Tone.Warn ? _theme.Warn : _theme.Accent;

            TextRenderer.DrawText(e.Graphics, _flashText, _font,
                new Rectangle(S(8), 0, Math.Max(0, ClientSize.Width - S(16)), ClientSize.Height), color,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter |
                TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix | TextFormatFlags.EndEllipsis);
        }

        private void Flash(string text, bool warn)
        {
            _flashText = text;
            _flashTone = warn ? Tone.Warn : Tone.Accent;
            _flashUntil = DateTime.UtcNow.AddSeconds(4);
            _flashActive = true;

            LayoutChips();   // hide the chips; the message is shown in their place
            Invalidate();
        }

        // ---------------- Theme ----------------

        /// <summary>
        /// Match the colors of the native status bar. The sample is taken from an area of the status bar
        /// not covered by this strip; if the sample is invalid, the previous theme is kept.
        /// </summary>
        public void ApplyTheme()
        {
            if (!Visible) return;

            // The theme almost never changes: checking every few seconds is enough.
            if (DateTime.UtcNow < _nextThemeCheck) return;
            _nextThemeCheck = DateTime.UtcNow.AddSeconds(2);

            Rectangle b = Bounds;
            Point pt = KeepNativeRightPx > 0
                ? new Point(b.Right + S(3), b.Top + b.Height / 2)
                : new Point(b.Left + b.Width / 2, b.Top - S(3));

            ThemeColors t = ThemeColors.TrySample(pt, _root);
            if (t == null || t.Back.ToArgb() == _theme.Back.ToArgb()) return;

            _theme = t;
            BackColor = t.Back;
            foreach (var chip in _chips.Values)
            {
                chip.Theme = t;
                chip.Invalidate();
            }
            Invalidate();
        }

        // ---------------- Data display ----------------

        private static string Fmt(double? v)
        {
            if (!v.HasValue) return "\u2013";
            return v.Value.ToString("#,0.##########", CultureInfo.CurrentCulture);
        }

        public void ApplySnapshot(SelectionSnapshot s)
        {
            if (s == null) return;   // read failed temporarily: keep the previous display

            foreach (var kind in StatInfo.Order)
            {
                var chip = _chips[kind];
                chip.SetValue(s.HasRange ? Fmt(s.Values[(int)kind]) : "\u2013");
                chip.Draggable = s.HasRange && s.CanBuildFormula;

                string tip = BuildTip(s, kind);
                string old;
                if (!_tips.TryGetValue(kind, out old) || old != tip)
                {
                    _tips[kind] = tip;
                    _tip.SetToolTip(chip, tip);
                }
            }

            LayoutChips();   // chip width may change with the length of the number
        }

        private static string BuildTip(SelectionSnapshot s, StatKind kind)
        {
            if (!s.HasRange) return StatInfo.Label(kind);
            if (!s.CanBuildFormula) return StatInfo.Label(kind) + "\nCannot be turned into a formula";
            if (s.Refs.Count == 0) return StatInfo.Label(kind) + "\nDrag onto a cell to create this formula";

            return s.BuildFormula(kind, s.ListSeparator) + "\nDrag onto a cell to create this formula";
        }

        // ---------------- Drag & drop chip -> cell ----------------

        private bool OnChipDragStarting(StatChip chip)
        {
            var app = Globals.Application;
            if (app == null) return false;

            // Re-read right now: the filter / selection may have changed since the last display.
            SelectionSnapshot snap = null;
            try { snap = StatEngine.Capture(app); }
            catch { }

            if (snap == null || !snap.HasRange || !snap.CanBuildFormula)
            {
                Flash("No range selection available to build a formula from", true);
                return false;
            }

            _dragSnap = snap;
            _dragKind = chip.Kind;
            _dragFormulaLocal = snap.BuildFormula(chip.Kind, snap.ListSeparator);
            DragActive = true;
            _lastLookup = 0;

            _ghost = new DragGhost(_theme);
            _ghost.SetContent(_dragFormulaLocal, null, GhostState.NoTarget);
            _ghost.MoveNear(Cursor.Position);
            _ghost.Show();
            return true;
        }

        private void OnChipDragMoved(StatChip chip)
        {
            if (_ghost == null) return;

            Point p = Cursor.Position;
            _ghost.MoveNear(p);

            // Look up the cell under the cursor at most every 50 ms.
            int now = Environment.TickCount;
            if (unchecked(now - _lastLookup) < 50) return;
            _lastLookup = now;

            var app = Globals.Application;
            if (app == null) return;

            Excel.Range target = StatEngine.RangeAt(app, p);
            try
            {
                if (target == null)
                    _ghost.SetContent(_dragFormulaLocal, null, GhostState.NoTarget);
                else if (StatEngine.IsCircular(app, target))
                    _ghost.SetContent(_dragFormulaLocal, null, GhostState.Circular);
                else
                    _ghost.SetContent(_dragFormulaLocal, StatEngine.AddressOf(target), GhostState.Valid);
            }
            finally
            {
                StatEngine.Release(target);
            }
        }

        private void OnChipDragEnded(StatChip chip, bool dropped)
        {
            var snap = _dragSnap;
            var kind = _dragKind;
            var app = Globals.Application;

            EndDrag();

            if (!dropped || snap == null || app == null) return;

            Excel.Range target = StatEngine.RangeAt(app, Cursor.Position);
            try
            {
                if (target == null)
                {
                    Flash("Drop onto a worksheet cell to create a formula", true);
                    return;
                }

                if (StatEngine.IsCircular(app, target))
                {
                    Flash("Target cell is inside the source range (circular reference)", true);
                    return;
                }

                string address = StatEngine.AddressOf(target);

                try
                {
                    // Do not use Range.Formula = ... here: an automation write clears
                    // Excel's native Undo history. Instead, send the formula through the clipboard and
                    // run Excel's Paste command. Paste is an Excel-native operation, so
                    // Excel itself creates the native Undo record.
                    string formula = snap.BuildFormula(kind, snap.ListSeparator);

                    if (!Commands.PasteFormulaNative(app, target, formula))
                    {
                        Flash("Failed to write formula at " + address, true);
                        return;
                    }

                    Flash("\u2713 " + address + "  " + formula, false);
                }
                catch (Exception ex)
                {
                    Flash("Failed to write formula at " + address + ": " + ex.Message.Trim(), true);
                }
            }
            finally
            {
                StatEngine.Release(target);
            }

            // Native Paste already scheduled one post-edit refresh through StatHost.
            // Do not enqueue another immediate refresh here.
        }

        private void EndDrag()
        {
            DragActive = false;
            _dragSnap = null;

            if (_ghost != null)
            {
                try { _ghost.Hide(); _ghost.Dispose(); }
                catch { }
                _ghost = null;
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                EndDrag();
                _tip.Dispose();
                _font.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
