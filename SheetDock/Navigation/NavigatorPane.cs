using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetDock
{
    // ============================================================
    //  ITEM
    // ============================================================
    public class SheetTabItem
    {
        public string Name { get; set; }
        public Color TabColor { get; set; } = Color.Empty;
        public bool IsSelected { get; set; }
        public bool IsHidden { get; set; }

        /// <summary>1-based index in wb.Sheets. Differs from the list index because hidden sheets are not displayed.</summary>
        public int ExcelIndex { get; set; }

        /// <summary>The sheet currently active in Excel (marked with a symbol in the pane).</summary>
        public bool IsActive { get; set; }
        public override string ToString() => Name;
    }

    // ============================================================
    //  LISTBOX
    // ============================================================
    public class SheetTabListBox : ListBox
    {
        private const int ItemPaddingX = 10;
        private const int ItemPaddingY = 4;
        public const int DragHandleWidth = 18;   // now on the right side of the item
        public const int MarkerWidth = 16;       // active-sheet marker (left)

        public Color ThemeBack { get; set; } = Color.FromArgb(240, 240, 240);
        public Color ThemeFore { get; set; } = Color.FromArgb(60, 60, 60);
        public Color ActiveBack { get; set; } = Color.White;
        public Color ActiveFore { get; set; } = Color.Black;
        public Color SelectedBack { get; set; } = Color.FromArgb(0, 120, 215);
        public Color SelectedFore { get; set; } = Color.White;
        public Color Divider { get; set; } = Color.FromArgb(210, 210, 210);
        public Color AccentLine { get; set; } = Color.FromArgb(0, 120, 215);
        public Color DragHandleColor { get; set; } = Color.FromArgb(160, 160, 160);
        public Color HiddenFore { get; set; } = Color.FromArgb(140, 140, 140);

        public int CurrentDragTarget { get; set; } = -1;
        public int HoverIndex { get { return _hoverIndex; } }

        /// <summary>
        /// When true, mouse-move is not forwarded to the native ListBox handler, so
        /// the ListBox doesn't drag-select (and repaint) the items passed over while dragging the handle.
        /// </summary>
        public bool SuppressNativeMouseSelect { get; set; }

        private int _lastTheme = -1;
        private int _hoverIndex = -1;
        private const int WM_MOUSEMOVE = 0x0200;

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_MOUSEMOVE && SuppressNativeMouseSelect)
            {
                long lp = m.LParam.ToInt64();
                int x = (short)(lp & 0xFFFF);
                int y = (short)((lp >> 16) & 0xFFFF);
                OnMouseMove(new MouseEventArgs(MouseButtons.Left, 0, x, y, 0));
                return;
            }

            base.WndProc(ref m);
        }

        /// <summary>Repaint a single item only, not the whole control.</summary>
        public void InvalidateItem(int index)
        {
            if (index < 0 || index >= Items.Count) return;
            Invalidate(GetItemRectangle(index));
        }

        public SheetTabListBox()
        {
            DrawMode = DrawMode.OwnerDrawFixed;
            ItemHeight = 30;
            BorderStyle = BorderStyle.None;
            IntegralHeight = false;
            Font = new Font("Segoe UI", 9f);
            MultiColumn = true;
            ColumnWidth = 170;
            HorizontalScrollbar = true;
            BackColor = ThemeBack;
            SelectionMode = SelectionMode.MultiExtended;
            AllowDrop = true;

            // ControlStyles only, no window style changes. WS_EX_COMPOSITED was tried and
            // failed ("Error creating window handle") because this control is hosted in an Excel CTP.
            DoubleBuffered = true;
        }

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hWnd);

        /// <summary>
        /// Follow the Office/Excel theme. The background color is read directly from Excel's pixels (the task
        /// pane header right above the pane), so it fits White / Colorful / Dark Gray / Black without
        /// depending on the registry. If the pane isn't shown yet, the registry is used as a temporary
        /// estimate; once the pane is shown, the theme is recomputed from pixels.
        /// </summary>
        public void ApplyThemeFromExcel()
        {
            try
            {
                if (Globals.Application == null) return;

                Color back;
                int key;

                Color? sampled = SampleExcelBackground();
                if (sampled.HasValue)
                {
                    back = sampled.Value;
                    key = back.ToArgb();
                }
                else
                {
                    int kind = GuessThemeKindFromRegistry();   // 0 light, 1 dark gray, 2 black
                    back = kind == 0 ? Color.White
                         : kind == 1 ? Color.FromArgb(68, 68, 68)
                         : Color.FromArgb(38, 38, 38);
                    key = -1 - kind;
                }

                // Theme unchanged -> don't set BackColor / Invalidate again (source of flicker).
                if (_lastTheme == key) return;
                _lastTheme = key;

                bool dark = (0.299 * back.R + 0.587 * back.G + 0.114 * back.B) < 128;

                ThemeBack = back;
                if (dark)
                {
                    ThemeFore = Color.FromArgb(242, 242, 242);
                    Divider = Blend(back, Color.White, 0.14);
                    DragHandleColor = Blend(back, Color.White, 0.40);
                    HiddenFore = Blend(back, Color.White, 0.45);
                    AccentLine = Color.FromArgb(16, 124, 65);
                }
                else
                {
                    ThemeFore = Color.FromArgb(38, 38, 38);
                    Divider = Blend(back, Color.Black, 0.12);
                    DragHandleColor = Blend(back, Color.Black, 0.30);
                    HiddenFore = Blend(back, Color.Black, 0.45);
                    AccentLine = Color.FromArgb(33, 115, 70);
                }

                SelectedBack = AccentLine;       // Excel green
                SelectedFore = Color.White;
                ActiveBack = ThemeBack;
                ActiveFore = ThemeFore;

                BackColor = ThemeBack;
                Invalidate();
            }
            catch { }
        }

        private static Color Blend(Color from, Color to, double t)
        {
            return Color.FromArgb(
                (int)(from.R + (to.R - from.R) * t),
                (int)(from.G + (to.G - from.G) * t),
                (int)(from.B + (to.B - from.B) * t));
        }

        /// <summary>Color of the pixel right above the pane (Excel's task pane header), or null.</summary>
        private Color? SampleExcelBackground()
        {
            try
            {
                if (!IsHandleCreated || Width < 20 || Height < 20) return null;
                if (!IsWindowVisible(Handle)) return null;

                Point pt = PointToScreen(new Point(Width / 2, -10));

                using (var bmp = new Bitmap(1, 1))
                using (var g = Graphics.FromImage(bmp))
                {
                    g.CopyFromScreen(pt, Point.Empty, new Size(1, 1));
                    return bmp.GetPixel(0, 0);
                }
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Estimate from the Office registry ("UI Theme": 0 Colorful, 3 Dark Gray, 4 Black, 5 White,
        /// 6 follow system). A missing / unrecognized value is treated as light.
        /// </summary>
        private static int GuessThemeKindFromRegistry()
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
                        if (i == 3) return 1;
                        if (i == 4) return 2;
                        if (i == 6) return IsWindowsDarkMode() ? 2 : 0;
                    }
                }
            }
            catch { }

            return 0;
        }

        private static bool IsWindowsDarkMode()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                    @"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"))
                {
                    if (key != null)
                    {
                        object val = key.GetValue("AppsUseLightTheme");
                        if (val is int i && i == 0) return true;
                    }
                }
            }
            catch { }
            return false;
        }

        /// <summary>
        /// Check whether x is within the drag handle area (right side of the item).
        /// </summary>
        public bool IsInDragHandle(int x)
        {
            int idx = IndexFromPoint(x, 1);
            if (idx < 0) return false;

            Rectangle r = GetItemRectangle(idx);
            return x >= r.Right - DragHandleWidth && x <= r.Right;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            int idx = IndexFromPoint(e.Location);
            if (idx >= Items.Count) idx = -1;
            if (idx != _hoverIndex)
            {
                int old = _hoverIndex;
                _hoverIndex = idx;
                InvalidateItem(old);
                InvalidateItem(_hoverIndex);
            }
        }

        protected override void OnMouseLeave(EventArgs e)
        {
            base.OnMouseLeave(e);
            if (_hoverIndex >= 0)
            {
                int old = _hoverIndex;
                _hoverIndex = -1;
                InvalidateItem(old);
            }
        }

        protected override void OnDrawItem(DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= Items.Count) return;

            var item = Items[e.Index] as SheetTabItem;
            string text = item?.Name ?? Items[e.Index]?.ToString() ?? string.Empty;
            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            bool isActive = item != null && item.IsActive;
            bool isHover = e.Index == _hoverIndex;
            bool hasCustomColor = item != null && item.TabColor != Color.Empty;

            // SheetDock uses the same palette as StatDock: the Excel background as the base,
            // then chip fill/hover/border/accent come from ThemeColors.
            var statTheme = StatDock.ThemeColors.FromBack(ThemeBack);
            Color fill = statTheme.ChipBack;
            Color fore = statTheme.Fore;
            Color border = statTheme.ChipBorder;

            if (hasCustomColor)
            {
                // The Excel tab color stays visible but is softened so the chip shape
                // stays consistent with StatDock.
                fill = StatDock.ThemeColors.Blend(ThemeBack, item.TabColor, 0.28);
                border = StatDock.ThemeColors.Blend(item.TabColor, fore, 0.35);
            }

            if (isHover && !isSelected)
                fill = statTheme.ChipHover;

            if (isSelected)
            {
                fill = statTheme.Accent;
                fore = Color.White;
                border = statTheme.Accent;
            }

            if (item != null && item.IsHidden)
                fore = isSelected ? Color.White : statTheme.Dim;

            // Clear the item area with the pane background first.
            using (var clear = new SolidBrush(ThemeBack))
                e.Graphics.FillRectangle(clear, e.Bounds);

            // Chip with inter-item margin, radius and border matching the visual character
            // of StatDock.StatChip.
            const int marginX = 4;
            const int marginY = 3;
            Rectangle chipRect = new Rectangle(
                e.Bounds.X + marginX,
                e.Bounds.Y + marginY,
                Math.Max(0, e.Bounds.Width - marginX * 2),
                Math.Max(0, e.Bounds.Height - marginY * 2));

            using (var path = RoundRect(chipRect, 5))
            using (var brush = new SolidBrush(fill))
            using (var pen = new Pen(border))
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                e.Graphics.FillPath(brush, path);
                e.Graphics.DrawPath(pen, path);
            }

            // Active sheet marker.
            if (isActive)
                DrawActiveMarker(e.Graphics, chipRect, isSelected ? Color.White : statTheme.Accent);

            int textLeft = chipRect.X + MarkerWidth + 1;
            int textRight = chipRect.Right - DragHandleWidth - 3;
            var textRect = new Rectangle(
                textLeft,
                chipRect.Y + ItemPaddingY - 1,
                Math.Max(0, textRight - textLeft),
                Math.Max(0, chipRect.Height - (ItemPaddingY * 2) + 2));

            TextRenderer.DrawText(e.Graphics, text, Font, textRect, fore,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);

            // The drag handle stays inside the chip.
            DrawDragHandle(e.Graphics, chipRect, item != null && item.IsHidden,
                isSelected ? (Color?)Color.FromArgb(210, Color.White) : null);

            // The drop target indicator follows the StatDock accent.
            if (CurrentDragTarget == e.Index)
            {
                using (var pen = new Pen(statTheme.Accent, 2f))
                {
                    e.Graphics.DrawLine(pen, chipRect.Left + 2, chipRect.Top + 1, chipRect.Left + 2, chipRect.Bottom - 1);
                }
            }
        }

        private static GraphicsPath RoundRect(Rectangle r, int radius)
        {
            int d = Math.Max(2, radius * 2);
            var p = new GraphicsPath();
            p.AddArc(r.X, r.Y, d, d, 180, 90);
            p.AddArc(r.Right - d, r.Y, d, d, 270, 90);
            p.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
            p.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
            p.CloseFigure();
            return p;
        }

        /// <summary>Small triangle (▶) marking the active sheet.</summary>
        private static void DrawActiveMarker(Graphics g, Rectangle bounds, Color color)
        {
            int cy = bounds.Top + bounds.Height / 2;
            int x0 = bounds.Left + 5;
            var pts = new[]
            {
                new Point(x0, cy - 4),
                new Point(x0 + 5, cy),
                new Point(x0, cy + 4)
            };

            var oldMode = g.SmoothingMode;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            using (var brush = new SolidBrush(color))
                g.FillPolygon(brush, pts);
            g.SmoothingMode = oldMode;
        }

        /// <summary>
        /// Draw the ⠿ drag handle on the right side of the item.
        /// The handle stays visible on hidden sheets, but with a dimmer color.
        /// </summary>
        private void DrawDragHandle(Graphics g, Rectangle bounds, bool isHidden, Color? tint = null)
        {
            Color color = tint ?? (isHidden
                ? Color.FromArgb(80, DragHandleColor)
                : DragHandleColor);

            using (var brush = new SolidBrush(color))
            {
                int dotSize = 2;
                int spacing = 3;
                int totalHeight = (dotSize + spacing) * 3 - spacing;
                int startY = bounds.Top + (bounds.Height - totalHeight) / 2;
                int x = bounds.Right - DragHandleWidth + 4;

                for (int row = 0; row < 3; row++)
                {
                    int y = startY + row * (dotSize + spacing);

                    g.FillEllipse(brush, x, y, dotSize, dotSize);
                    g.FillEllipse(brush, x + dotSize + 2, y, dotSize, dotSize);
                }
            }
        }

        private static Color GetInverseTextColor(Color bg)
        {
            double lum = (0.299 * bg.R + 0.587 * bg.G + 0.114 * bg.B);
            return lum < 128 ? Color.White : Color.Black;
        }
    }

    // ============================================================
    //  NAVIGATOR PANE
    // ============================================================
    public class NavigatorPane : IDisposable
    {
        private readonly Panel _host;
        private readonly SheetTabListBox _sheetList;
        private readonly ContextMenuStrip _contextMenu;
        private bool _disposed;

        // Drag state
        private bool _isDragging;
        private int _dragIndex = -1;
        private int _dragTarget = -1;
        private Point _dragStartPoint;
        private readonly Timer _autoScrollTimer;
        private readonly Timer _themeTimer;   // recompute the theme after the pane is shown / resized
        private int _autoScrollDir;

        // > 0 while we change the selection programmatically (refresh, etc.),
        // so SelectedIndexChanged doesn't trigger ws.Activate() in Excel.
        private int _suppressActivate;
        // A mouse click changes the selection while the button is still down; activation is deferred to MouseUp.
        private bool _activatePending;
        // ---- Drag-to-select (rubber band) ----
        private bool _bandActive;
        private bool _bandStarted;
        private Point _bandOffset;   // start point, relative to item 0 (scroll-proof)
        private readonly HashSet<int> _bandBase = new HashSet<int>();

        private string _listWorkbookName;
        private string _lastActiveName;         // active sheet at the last refresh/sync   // workbook that owns the current list contents

        // ---- Registry of all panes (one per Excel window) ----
        private static readonly List<NavigatorPane> _instances = new List<NavigatorPane>();
        private static Timer _refreshTimer;

        /// <summary>The pane that is showing the context menu (target of the PanelCommands macros).</summary>
        public static NavigatorPane ContextPane;

        /// <summary>Refresh all live panes (one per CTP).</summary>
        public static void RefreshAll()
        {
            foreach (var p in _instances.ToArray())
            {
                try { p.RefreshSheetList(); } catch { }
            }
        }

        /// <summary>
        /// Coalesced version of RefreshAll: WorkbookActivate + WindowActivate + WorkbookOpen
        /// usually fire back-to-back, so they are merged into a single refresh.
        /// </summary>
        public static void RequestRefreshAll()
        {
            if (_refreshTimer == null)
            {
                _refreshTimer = new Timer { Interval = 40 };
                _refreshTimer.Tick += (s, e) =>
                {
                    _refreshTimer.Stop();
                    RefreshAll();
                };
            }
            _refreshTimer.Stop();
            _refreshTimer.Start();
        }

        /// <summary>Sync the active-sheet marker across all panes (called from SheetActivate).</summary>
        public static void SyncAllActive()
        {
            foreach (var p in _instances.ToArray())
            {
                try { p.SyncActiveSheet(); } catch { }
            }
        }

        public static void ShutdownStatics()
        {
            try { _refreshTimer?.Stop(); _refreshTimer?.Dispose(); } catch { }
            _refreshTimer = null;
        }

        public NavigatorPane()
        {
            _host = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Color.FromArgb(240, 240, 240)
            };

            _sheetList = new SheetTabListBox { Dock = DockStyle.Fill };
            _sheetList.DoubleClick += SheetList_DoubleClick;
            _sheetList.SelectedIndexChanged += SheetList_SelectedIndexChanged;
            _sheetList.MouseDown += SheetList_MouseDown;
            _sheetList.MouseMove += SheetList_MouseMove;
            _sheetList.MouseUp += SheetList_MouseUp;
            _sheetList.KeyDown += SheetList_KeyDown;
            _sheetList.DragEnter += (s, e) => e.Effect = DragDropEffects.Move;
            _sheetList.DragDrop += (s, e) => EndDrag();

            _host.Controls.Add(_sheetList);

            // Fallback menu (used if Excel's native menu fails to show).
            // Deliberately not assigned to ContextMenuStrip: right-click is handled by ShowSheetMenu.
            _contextMenu = BuildContextMenu();

            _autoScrollTimer = new Timer { Interval = 60 };
            _autoScrollTimer.Tick += AutoScrollTimer_Tick;

            _themeTimer = new Timer { Interval = 250 };
            _themeTimer.Tick += (s, e) =>
            {
                _themeTimer.Stop();
                _sheetList.ApplyThemeFromExcel();
                _host.BackColor = _sheetList.ThemeBack;
            };
            _host.VisibleChanged += (s, e) => { _themeTimer.Stop(); _themeTimer.Start(); };
            _sheetList.SizeChanged += (s, e) => { _themeTimer.Stop(); _themeTimer.Start(); };

            _sheetList.ApplyThemeFromExcel();
            RefreshSheetList();

            _instances.Add(this);
        }

        public Control HostControl => _host;

        // ========================================================
        //  CONTEXT MENU
        // ========================================================
        private ContextMenuStrip BuildContextMenu()
        {
            var menu = new ContextMenuStrip();

            menu.Items.Add("Add", null, (s, e) => AddSheet());
            menu.Items.Add("Insert From File...", null, (s, e) => InsertFromFile());
            menu.Items.Add("Duplicate", null, (s, e) => DuplicateSheet());
            menu.Items.Add("Duplicate Selection", null, (s, e) => DuplicateSelection());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Rename", null, (s, e) => RenameSelectedSheet());
            menu.Items.Add("Change Color", null, (s, e) => ChangeTabColor());
            menu.Items.Add("Remove Color", null, (s, e) => ClearTabColor());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Hide", null, (s, e) => HideSelectedSheet());
            menu.Items.Add("Unhide...", null, (s, e) => UnhideSheetsDialog());
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Select All Sheets", null, (s, e) => SelectAllSheets());
            menu.Items.Add("Ungroup Sheets", null, (s, e) => UngroupSheets());

            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("Delete", null, (s, e) => DeleteSelectedSheet());

            return menu;
        }

        // ========================================================
        //  NATIVE CONTEXT MENU (CommandBar "Ply" = Excel's sheet tab right-click menu)
        // ========================================================
        public void CmdAddSheet() => AddSheet();
        public void CmdInsertFromFile() => InsertFromFile();
        public void CmdDuplicate() => DuplicateSheet();
        public void CmdDuplicateSelection() => DuplicateSelection();

        private void ShowSheetMenu(Point clientPoint)
        {
            if (_sheetList.SelectedIndices.Count == 0) return;

            try
            {
                ShowNativeSheetMenu();
            }
            catch
            {
                _contextMenu.Show(_sheetList, clientPoint);
            }
        }

        private void ShowNativeSheetMenu()
        {
            var app = Globals.Application;
            var wb = ActiveWorkbook();
            if (app == null || wb == null) throw new InvalidOperationException("No workbook");

            var names = SelectedSheetNames();
            if (names.Count == 0) return;

            // The native menu operates on the sheets selected in Excel, so first sync
            // with the selection in the pane (multiple sheets = temporary group).
            bool grouped = names.Count > 1;
            SyncExcelSelection(wb, names);

            var bar = app.CommandBars["Ply"];
            var added = new List<Office.CommandBarControl>();

            ContextPane = this;
            try
            {
                // Extra items, installed temporarily on top of the native menu.
                AddPlyItem(bar, "Add Sheet", "SheetDock_AddSheet", 1, added);
                AddPlyItem(bar, "Insert From File...", "SheetDock_InsertFromFile", 2, added);
                AddPlyItem(bar, "Duplicate", "SheetDock_Duplicate", 3, added);
                AddPlyItem(bar, "Duplicate Selection", "SheetDock_DuplicateSelection", 4, added);

                bar.ShowPopup(Type.Missing, Type.Missing);
            }
            finally
            {
                foreach (var c in added)
                {
                    try { c.Delete(true); } catch { }
                }

                // Runs after the menu command finishes: dissolve the temporary group
                // (so subsequent edits don't overwrite multiple sheets) and sync the pane.
                string keep = names[0];
                int groupSize = names.Count;
                ExcelDna.Integration.ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        if (grouped)
                        {
                            var sel = app.ActiveWindow?.SelectedSheets;
                            if (sel != null && sel.Count == groupSize)
                                ActivateSheetByName(keep);
                        }
                    }
                    catch { }

                    RefreshAll();
                });
            }
        }

        private static void AddPlyItem(Office.CommandBar bar, string caption, string macro,
            int position, List<Office.CommandBarControl> added)
        {
            var ctl = bar.Controls.Add(Office.MsoControlType.msoControlButton,
                Type.Missing, Type.Missing, position, true);

            var btn = (Office.CommandBarButton)ctl;
            btn.Caption = caption;
            btn.OnAction = macro;
            btn.Style = Office.MsoButtonStyle.msoButtonCaption;

            added.Add(ctl);
        }

        private void SyncExcelSelection(Excel.Workbook wb, List<string> names)
        {
            if (names.Count == 1)
            {
                ActivateSheetByName(names[0]);
                return;
            }

            try
            {
                object[] arr = names.Cast<object>().ToArray();
                var sheets = wb.Sheets[arr] as Excel.Sheets;
                sheets?.Select(Type.Missing);
            }
            catch
            {
                ActivateSheetByName(names[0]);
            }
        }

        // ========================================================
        //  DRAG & DROP (⠿ drag handle only)
        // ========================================================
        private void SheetList_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int idx = _sheetList.IndexFromPoint(e.Location);
                if (idx >= 0 && idx < _sheetList.Items.Count)
                {
                    if (!_sheetList.SelectedIndices.Contains(idx))
                    {
                        _sheetList.ClearSelected();
                        _sheetList.SelectedIndex = idx;
                    }
                }
                return;
            }

            if (e.Button == MouseButtons.Left)
            {
                int idx = _sheetList.IndexFromPoint(e.Location);
                var item = idx >= 0 ? _sheetList.Items[idx] as SheetTabItem : null;

                if (item != null && !item.IsHidden)
                {
                    // Drag-reorder only starts when clicking in the ⠿ handle area (right)
                    Rectangle r = _sheetList.GetItemRectangle(idx);
                    if (e.X >= r.Right - SheetTabListBox.DragHandleWidth && e.X <= r.Right)
                    {
                        _dragIndex = idx;
                        _dragStartPoint = e.Location;
                        _isDragging = false;
                        _sheetList.SuppressNativeMouseSelect = true;
                        return;
                    }
                }

                _dragIndex = -1;
                BeginBand(e);
            }
        }

        // ========================================================
        //  DRAG-TO-SELECT
        // ========================================================
        // Hold the mouse on a row (or empty area) and drag: every item touched by
        // the drag rectangle gets selected. Ctrl = add to the existing selection.
        // Shift+click is still handled natively (range).
        private void BeginBand(MouseEventArgs e)
        {
            _bandActive = false;
            _bandStarted = false;

            if (_sheetList.Items.Count == 0) return;
            if ((Control.ModifierKeys & Keys.Shift) != 0) return;

            Rectangle r0 = _sheetList.GetItemRectangle(0);
            _bandOffset = new Point(e.X - r0.X, e.Y - r0.Y);

            _bandBase.Clear();
            if ((Control.ModifierKeys & Keys.Control) != 0)
            {
                foreach (int i in _sheetList.SelectedIndices)
                    _bandBase.Add(i);
            }

            _bandActive = true;
        }

        private void UpdateBand(Point p)
        {
            if (_sheetList.Items.Count == 0) return;

            Rectangle r0 = _sheetList.GetItemRectangle(0);
            var start = new Point(r0.X + _bandOffset.X, r0.Y + _bandOffset.Y);

            if (!_bandStarted)
            {
                if (Math.Abs(p.X - start.X) < 4 && Math.Abs(p.Y - start.Y) < 4)
                    return;

                _bandStarted = true;
                // From here on the selection is fully controlled by us, not the native ListBox.
                _sheetList.SuppressNativeMouseSelect = true;
            }

            var rect = Rectangle.FromLTRB(
                Math.Min(start.X, p.X), Math.Min(start.Y, p.Y),
                Math.Max(start.X, p.X) + 1, Math.Max(start.Y, p.Y) + 1);

            var result = new HashSet<int>(_bandBase);
            for (int i = 0; i < _sheetList.Items.Count; i++)
            {
                if (_sheetList.GetItemRectangle(i).IntersectsWith(rect))
                    result.Add(i);
            }

            if (result.SetEquals(_sheetList.SelectedIndices.Cast<int>()))
            {
                CheckAutoScroll(p);
                return;
            }

            _suppressActivate++;
            _sheetList.BeginUpdate();
            try
            {
                _sheetList.ClearSelected();
                foreach (int i in result)
                    _sheetList.SetSelected(i, true);
            }
            finally
            {
                _sheetList.EndUpdate();
                _suppressActivate--;
            }

            CheckAutoScroll(p);
        }

        private void SheetList_MouseMove(object sender, MouseEventArgs e)
        {
            if (_bandActive && e.Button == MouseButtons.Left)
            {
                UpdateBand(e.Location);
                return;
            }

            if (e.Button != MouseButtons.Left || _dragIndex < 0) return;

            // Only start dragging after movement > 4px
            if (!_isDragging)
            {
                if (Math.Abs(e.X - _dragStartPoint.X) < 4 &&
                    Math.Abs(e.Y - _dragStartPoint.Y) < 4)
                    return;

                _isDragging = true;
            }

            int target = _sheetList.IndexFromPoint(e.Location);
            if (target < 0) target = _sheetList.Items.Count - 1;

            if (target != _dragTarget)
            {
                int old = _dragTarget;
                _dragTarget = target;
                _sheetList.CurrentDragTarget = target;
                _sheetList.InvalidateItem(old);
                _sheetList.InvalidateItem(target);
            }

            // Auto-scroll when near the edge
            CheckAutoScroll(e.Location);
        }

        private void SheetList_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                ShowSheetMenu(e.Location);
                return;
            }

            if (e.Button != MouseButtons.Left) return;

            bool moved = false;

            if (_isDragging && _dragIndex >= 0 && _dragTarget >= 0 &&
                _dragTarget != _dragIndex)
            {
                moved = true;
                MoveAndReselect(_dragIndex, _dragTarget);
            }

            bool activate = _activatePending && !moved;
            _activatePending = false;

            EndDrag();

            // Plain click (including a click on the handle without dragging) -> activate the sheet on mouse release.
            if (activate && _sheetList.SelectedIndices.Count == 1)
                ActivateSelectedSheet();
        }

        /// <summary>
        /// Move the sheet, refresh the list, then re-select the moved sheet —
        /// without triggering sheet activation and with ScreenUpdating turned off
        /// so Excel doesn't repaint at every step.
        /// </summary>
        private void MoveAndReselect(int fromIdx, int toIdx)
        {
            var app = Globals.Application;
            bool oldScreenUpdating = true;
            try { if (app != null) { oldScreenUpdating = app.ScreenUpdating; app.ScreenUpdating = false; } }
            catch { }

            _suppressActivate++;
            try
            {
                string movedName = (_sheetList.Items[fromIdx] as SheetTabItem)?.Name;

                MoveSheet(fromIdx, toIdx);

                // The moved sheet becomes the active sheet (activation is deferred when pressing on the handle).
                if (movedName != null) ActivateSheetByName(movedName);

                RefreshSheetList();
            }
            catch { }
            finally
            {
                _suppressActivate--;
                try { if (app != null) app.ScreenUpdating = oldScreenUpdating; } catch { }
            }
        }

        private void EndDrag()
        {
            bool wasDragging = _isDragging;
            _isDragging = false;
            _bandActive = false;
            _bandStarted = false;
            _dragIndex = -1;
            _dragTarget = -1;
            _sheetList.CurrentDragTarget = -1;
            _sheetList.SuppressNativeMouseSelect = false;
            _autoScrollTimer.Stop();
            _autoScrollDir = 0;
            if (wasDragging) _sheetList.Invalidate();
        }

        private void CheckAutoScroll(Point p)
        {
            int margin = 20;

            // Pane docked top/bottom = a single row of tabs -> scroll follows the X axis.
            bool singleRow = _sheetList.ClientSize.Height < _sheetList.ItemHeight * 2;
            int pos = singleRow ? p.X : p.Y;
            int size = singleRow ? _sheetList.ClientSize.Width : _sheetList.ClientSize.Height;

            if (pos < margin)
            {
                _autoScrollDir = -1;
                _autoScrollTimer.Start();
            }
            else if (pos > size - margin)
            {
                _autoScrollDir = 1;
                _autoScrollTimer.Start();
            }
            else
            {
                _autoScrollTimer.Stop();
                _autoScrollDir = 0;
            }
        }

        private void AutoScrollTimer_Tick(object sender, EventArgs e)
        {
            if (_autoScrollDir == 0) return;

            try
            {
                int newTop = _sheetList.TopIndex + _autoScrollDir;
                if (newTop < 0) newTop = 0;
                if (newTop >= _sheetList.Items.Count)
                    newTop = _sheetList.Items.Count - 1;

                _sheetList.TopIndex = newTop;

                // Currently drag-to-selecting: recompute the rectangle after the list scrolls.
                if (_bandActive && _bandStarted)
                    UpdateBand(_sheetList.PointToClient(Cursor.Position));
            }
            catch { }
        }

        private void MoveSheet(int fromIdx, int toIdx)
        {
            try
            {
                var app = Globals.Application;
                if (app == null) return;

                var wb = app.ActiveWorkbook;
                if (wb == null) return;

                var fromItem = _sheetList.Items[fromIdx] as SheetTabItem;
                if (fromItem == null) return;

                var toItem = _sheetList.Items[toIdx] as SheetTabItem;
                if (toItem == null) return;

                object sheet = wb.Sheets[fromItem.Name];

                // Use ExcelIndex (not the list index): hidden sheets aren't in the list,
                // so list position != position in the workbook.
                object target = wb.Sheets[toItem.ExcelIndex];

                if (toIdx > fromIdx)
                {
                    if (sheet is Excel.Worksheet ws) ws.Move(After: target);
                    else if (sheet is Excel.Chart ch) ch.Move(After: target);
                }
                else
                {
                    if (sheet is Excel.Worksheet ws) ws.Move(Before: target);
                    else if (sheet is Excel.Chart ch) ch.Move(Before: target);
                }
            }
            catch { }
        }

        private void SheetList_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                DeleteSelectedSheet();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.F2)
            {
                RenameSelectedSheet();
                e.Handled = true;
            }
        }

        // ========================================================
        //  REFRESH
        // ========================================================
        public void RefreshSheetList()
        {
            if (_disposed) return;

            var selectedNames = new HashSet<string>();
            foreach (int i in _sheetList.SelectedIndices)
            {
                if (_sheetList.Items[i] is SheetTabItem it)
                    selectedNames.Add(it.Name);
            }

            // Don't carry the previous workbook's selection over to another workbook
            // (the same sheet name can exist in both).
            string wbName = null;
            try { wbName = Globals.Application?.ActiveWorkbook?.Name; } catch { }
            bool wbChanged = wbName != _listWorkbookName;
            if (wbChanged)
                selectedNames.Clear();
            _listWorkbookName = wbName;

            if (selectedNames.Count == 0)
            {
                try
                {
                    var app0 = Globals.Application;
                    if (app0?.ActiveSheet is Excel.Worksheet aws)
                        selectedNames.Add(aws.Name);
                    else if (app0?.ActiveSheet is Excel.Chart ach)
                        selectedNames.Add(ach.Name);
                }
                catch { }
            }

            int prevTop = _sheetList.TopIndex;

            _suppressActivate++;
            _sheetList.BeginUpdate();
            try
            {
                _sheetList.Items.Clear();

                var app = Globals.Application;
                if (app == null)
                {
                    _sheetList.Items.Add(new SheetTabItem { Name = "(Excel not available)" });
                    return;
                }

                Excel.Workbook wb = null;
                try { wb = app.ActiveWorkbook; } catch { }
                if (wb == null)
                {
                    _sheetList.Items.Add(new SheetTabItem { Name = "(no workbook)" });
                    return;
                }

                int count = 0;
                try { count = wb.Sheets.Count; } catch { }

                string activeName = GetActiveSheetName();

                // Iterate by index — covers hidden & very hidden sheets
                for (int i = 1; i <= count; i++)
                {
                    object sheetObj = null;
                    try { sheetObj = wb.Sheets[i]; } catch { }
                    if (sheetObj == null) continue;

                    // Hidden / very hidden sheets are not shown in the pane.
                    if (IsSheetHidden(sheetObj)) continue;

                    var item = BuildSheetItem(sheetObj);
                    if (item == null) continue;

                    item.ExcelIndex = i;
                    item.IsSelected = selectedNames.Contains(item.Name);
                    item.IsActive = item.Name == activeName;

                    _sheetList.Items.Add(item);
                }

                for (int i = 0; i < _sheetList.Items.Count; i++)
                {
                    if (_sheetList.Items[i] is SheetTabItem it && it.IsSelected)
                        _sheetList.SetSelected(i, true);
                }

                // Selection follows the active sheet (e.g. the old sheet was hidden, or the workbook changed).
                EnsureSelectionFollowsActive(activeName);

                // Items.Clear() resets the scroll to the top. If the active sheet changed (or the
                // workbook changed) scroll to the active sheet; otherwise keep the user's scroll position.
                if (wbChanged || activeName != _lastActiveName)
                {
                    EnsureItemVisible(IndexOfName(activeName));
                }
                else if (_sheetList.Items.Count > 0)
                {
                    _sheetList.TopIndex = Math.Max(0, Math.Min(prevTop, _sheetList.Items.Count - 1));
                }
                _lastActiveName = activeName;
            }
            finally
            {
                _sheetList.EndUpdate();
                _suppressActivate--;
            }

            _sheetList.ApplyThemeFromExcel();
            _host.BackColor = _sheetList.ThemeBack;
        }

        /// <summary>
        /// Scroll the list (vertical / horizontal, depending on the column layout) until the item is fully visible.
        /// Computed from the index, not relying on GetItemRectangle, which is unreliable
        /// for items outside the visible area.
        /// </summary>
        private void EnsureItemVisible(int index)
        {
            if (index < 0 || index >= _sheetList.Items.Count) return;

            int h = _sheetList.ClientSize.Height;
            int w = _sheetList.ClientSize.Width;
            if (h <= 0 || w <= 0 || _sheetList.ItemHeight <= 0) return;

            int rows = Math.Max(1, h / _sheetList.ItemHeight);
            int top = _sheetList.TopIndex;

            try
            {
                if (!_sheetList.MultiColumn)
                {
                    if (index < top) _sheetList.TopIndex = index;
                    else if (index >= top + rows) _sheetList.TopIndex = index - rows + 1;
                    return;
                }

                int cols = Math.Max(1, w / Math.Max(1, _sheetList.ColumnWidth));
                if (index >= top && index < top + rows * cols) return;

                int colStart = (index / rows) * rows;
                _sheetList.TopIndex = index < top
                    ? colStart
                    : Math.Max(0, colStart - (cols - 1) * rows);
            }
            catch { }
        }

        private static string GetActiveSheetName()
        {
            try
            {
                var app = Globals.Application;
                if (app?.ActiveSheet is Excel.Worksheet aws) return aws.Name;
                if (app?.ActiveSheet is Excel.Chart ach) return ach.Name;
            }
            catch { }
            return null;
        }

        private int IndexOfName(string name)
        {
            for (int i = 0; i < _sheetList.Items.Count; i++)
            {
                if (_sheetList.Items[i] is SheetTabItem it && it.Name == name)
                    return i;
            }
            return -1;
        }

        /// <summary>
        /// Ensure the list selection = Excel's active sheet. If the user is selecting
        /// several sheets that include the active sheet (a group), the selection is left alone.
        /// </summary>
        private void EnsureSelectionFollowsActive(string activeName)
        {
            if (activeName == null) return;

            int ai = IndexOfName(activeName);
            if (ai < 0) return;

            if (_sheetList.GetSelected(ai)) return;

            _suppressActivate++;
            try
            {
                _sheetList.ClearSelected();
                _sheetList.SetSelected(ai, true);
            }
            finally
            {
                _suppressActivate--;
            }
        }

        /// <summary>
        /// Update the active marker + selection without rebuilding the list. Called when a sheet
        /// is activated from outside the pane (Ctrl+[, hyperlink click, built-in tab, macro, etc.).
        /// </summary>
        public void SyncActiveSheet()
        {
            if (_disposed || _isDragging) return;

            string activeName = GetActiveSheetName();
            if (activeName == null) return;

            int ai = IndexOfName(activeName);
            if (ai < 0)
            {
                RefreshSheetList();   // new sheet / list out of sync (Refresh already scrolls to the active one)
                return;
            }

            bool changed = false;
            foreach (var it in _sheetList.Items.OfType<SheetTabItem>())
            {
                bool a = it.Name == activeName;
                if (it.IsActive != a) { it.IsActive = a; changed = true; }
            }

            EnsureSelectionFollowsActive(activeName);

            // Active sheet changed from outside the pane (Ctrl+[, hyperlink, etc.): scroll to it.
            EnsureItemVisible(ai);
            _lastActiveName = activeName;

            if (changed) _sheetList.Invalidate();
        }

        private static SheetTabItem BuildSheetItem(object sheetObj)
        {
            if (sheetObj is Excel.Worksheet ws)
            {
                return new SheetTabItem
                {
                    Name = ws.Name,
                    TabColor = ReadTabColor(ws.Tab)
                };
            }

            if (sheetObj is Excel.Chart ch)
            {
                return new SheetTabItem
                {
                    Name = ch.Name,
                    TabColor = ReadTabColor(ch.Tab)
                };
            }

            return null;
        }

        private static bool IsSheetHidden(object sheetObj)
        {
            try
            {
                if (sheetObj is Excel.Worksheet ws)
                    return ws.Visible != Excel.XlSheetVisibility.xlSheetVisible;

                if (sheetObj is Excel.Chart ch)
                    return ch.Visible != Excel.XlSheetVisibility.xlSheetVisible;
            }
            catch { }
            return false;
        }

        private static Color ReadTabColor(object tab)
        {
            if (tab == null) return Color.Empty;
            try
            {
                if (tab is Excel.Tab excelTab)
                {
                    object colorObj = excelTab.Color;
                    if (colorObj == null) return Color.Empty;
                    int ole = Convert.ToInt32(colorObj);
                    if (ole < 0) return Color.Empty;
                    return ColorTranslator.FromOle(ole);
                }
            }
            catch { return Color.Empty; }
            return Color.Empty;
        }

        // ========================================================
        //  EVENTS
        // ========================================================
        private void SheetList_DoubleClick(object sender, EventArgs e)
        {
            ActivateSelectedSheet();
        }

        private void SheetList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_suppressActivate > 0) return;
            if (_sheetList.SelectedIndices.Count != 1) return;

            // Selection changed by a mouse click (button still down): defer to MouseUp,
            // so pressing on the drag handle doesn't immediately activate the sheet in Excel.
            if (Control.MouseButtons == MouseButtons.Left)
            {
                _activatePending = true;
                return;
            }

            ActivateSelectedSheet();
        }

        private void ActivateSelectedSheet()
        {
            var item = _sheetList.SelectedItem as SheetTabItem;
            if (item == null || item.IsHidden) return;

            ActivateSheetByName(item.Name);
        }

        private void ActivateSheetByName(string name)
        {
            var app = Globals.Application;
            if (app == null) return;

            Excel.Workbook wb = null;
            try { wb = app.ActiveWorkbook; } catch { }
            if (wb == null) return;

            try
            {
                object sheet = wb.Sheets[name];

                if (sheet is Excel.Worksheet ws)
                {
                    ws.Activate();
                    Globals.Recent.Push(ws.Name);
                }
                else if (sheet is Excel.Chart ch)
                {
                    ch.Activate();
                    Globals.Recent.Push(ch.Name);
                }
            }
            catch { }
        }

        // ========================================================
        //  ACTIONS
        // ========================================================
        private Excel.Workbook ActiveWorkbook()
        {
            try { return Globals.Application?.ActiveWorkbook; }
            catch { return null; }
        }

        private void AddSheet()
        {
            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                object after = wb.ActiveSheet;
                var newSheet = (Excel.Worksheet)wb.Worksheets.Add(
                    Type.Missing, after, Type.Missing, Type.Missing);

                RefreshSheetList();

                for (int i = 0; i < _sheetList.Items.Count; i++)
                {
                    if (_sheetList.Items[i] is SheetTabItem it &&
                        it.Name == newSheet.Name)
                    {
                        _sheetList.ClearSelected();
                        _sheetList.SelectedIndex = i;
                        break;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to add sheet:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void InsertFromFile()
        {
            var wb = ActiveWorkbook();
            if (wb == null) return;

            using (var dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel Files|*.xlsx;*.xlsm;*.xls;*.xlsb|All Files|*.*";
                dlg.Title = "Select an Excel file to insert from";

                if (dlg.ShowDialog() != DialogResult.OK) return;

                Excel.Workbook srcWb = null;
                try
                {
                    srcWb = Globals.Application.Workbooks.Open(dlg.FileName, ReadOnly: true);

                    using (var picker = new SheetPickerForm(srcWb))
                    {
                        if (picker.ShowDialog() != DialogResult.OK)
                        {
                            srcWb.Close(SaveChanges: false);
                            return;
                        }

                        foreach (var sheetName in picker.SelectedSheetNames)
                        {
                            var srcSheet = (Excel.Worksheet)srcWb.Worksheets[sheetName];
                            srcSheet.Copy(After: wb.Sheets[wb.Sheets.Count]);
                        }
                    }

                    srcWb.Close(SaveChanges: false);
                    RefreshSheetList();
                }
                catch (Exception ex)
                {
                    try { srcWb?.Close(SaveChanges: false); } catch { }
                    MessageBox.Show("Failed to insert sheet:\n" + ex.Message,
                        "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void DuplicateSheet()
        {
            var item = _sheetList.SelectedItem as SheetTabItem;
            if (item == null) return;

            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                // Excel refuses to copy sheets that contain tables while the sheets are grouped.
                // Make sure only this sheet is selected before Copy.
                SelectOnly(wb, item.Name);

                object sheet = wb.Sheets[item.Name];
                if (sheet is Excel.Worksheet ws)
                {
                    ws.Copy(After: ws);
                    RefreshSheetList();
                }
                else
                {
                    MessageBox.Show("Duplicate is only available for worksheets.",
                        "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to duplicate:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DuplicateSelection()
        {
            var wb = ActiveWorkbook();
            if (wb == null) return;

            var selected = SelectedSheetNames();

            if (selected.Count == 0) return;

            var app = Globals.Application;
            bool oldScreenUpdating = true;
            try { if (app != null) { oldScreenUpdating = app.ScreenUpdating; app.ScreenUpdating = false; } }
            catch { }

            try
            {
                foreach (var name in selected)
                {
                    // One at a time, and each time only that sheet is selected
                    // (copying a sheet with tables fails while sheets are grouped).
                    SelectOnly(wb, name);

                    object sheet = wb.Sheets[name];

                    if (sheet is Excel.Worksheet ws)
                        ws.Copy(After: wb.Sheets[wb.Sheets.Count]);
                    else if (sheet is Excel.Chart ch)
                        ch.Copy(After: wb.Sheets[wb.Sheets.Count]);
                }

                RefreshSheetList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to duplicate selection:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                try { if (app != null) app.ScreenUpdating = oldScreenUpdating; } catch { }
            }
        }

        /// <summary>Select only one sheet in Excel (dissolves the sheet group).</summary>
        private static void SelectOnly(Excel.Workbook wb, string name)
        {
            object sheet = wb.Sheets[name];
            if (sheet is Excel.Worksheet ws) ws.Select(Type.Missing);
            else if (sheet is Excel.Chart ch) ch.Select(Type.Missing);
        }

        private void RenameSelectedSheet()
        {
            var item = _sheetList.SelectedItem as SheetTabItem;
            if (item == null) return;

            string oldName = item.Name;
            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                object sheet = wb.Sheets[oldName];
                string newName = Microsoft.VisualBasic.Interaction.InputBox(
                    "New name for the sheet:", "Rename Sheet", oldName);

                if (string.IsNullOrWhiteSpace(newName) || newName == oldName) return;

                if (sheet is Excel.Worksheet ws) ws.Name = newName;
                else if (sheet is Excel.Chart ch) ch.Name = newName;

                RefreshSheetList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to rename:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ChangeTabColor()
        {
            var names = SelectedSheetNames();
            if (names.Count == 0) return;

            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                // Prefill the dialog with the color of the first selected sheet, so
                // reopening the color picker on an already-colored group feels natural.
                var firstItem = _sheetList.Items.OfType<SheetTabItem>()
                    .FirstOrDefault(it => it.Name == names[0]);

                using (var dlg = new ColorDialog())
                {
                    dlg.FullOpen = true;
                    if (firstItem != null && firstItem.TabColor != Color.Empty)
                        dlg.Color = firstItem.TabColor;
                    if (dlg.ShowDialog() != DialogResult.OK) return;

                    int oleColor = ColorTranslator.ToOle(dlg.Color);

                    foreach (var name in names)
                    {
                        object sheet = wb.Sheets[name];
                        if (sheet is Excel.Worksheet ws) ws.Tab.Color = oleColor;
                        else if (sheet is Excel.Chart ch) ch.Tab.Color = oleColor;
                    }

                    RefreshSheetList();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to change color:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ClearTabColor()
        {
            var names = SelectedSheetNames();
            if (names.Count == 0) return;

            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                foreach (var name in names)
                {
                    object sheet = wb.Sheets[name];
                    if (sheet is Excel.Worksheet ws)
                        ws.Tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                    else if (sheet is Excel.Chart ch)
                        ch.Tab.ColorIndex = Excel.XlColorIndex.xlColorIndexNone;
                }
                RefreshSheetList();
            }
            catch { }
        }

        /// <summary>Names of all sheets currently selected in the pane (list order).</summary>
        private List<string> SelectedSheetNames()
        {
            var names = new List<string>();
            foreach (int i in _sheetList.SelectedIndices)
                if (_sheetList.Items[i] is SheetTabItem it)
                    names.Add(it.Name);
            return names;
        }

        private void HideSelectedSheet()
        {
            var wb = ActiveWorkbook();
            if (wb == null) return;

            try
            {
                foreach (var name in SelectedSheetNames())
                {
                    object sheet = wb.Sheets[name];
                    if (sheet is Excel.Worksheet ws)
                        ws.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                    else if (sheet is Excel.Chart ch)
                        ch.Visible = Excel.XlSheetVisibility.xlSheetHidden;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to hide:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Refresh even if some failed (e.g. Excel refuses to hide the last visible sheet).
                RefreshSheetList();
            }
        }

        /// <summary>Excel's built-in Unhide dialog.</summary>
        private void UnhideSheetsDialog()
        {
            var app = Globals.Application;
            if (app == null) return;

            try
            {
                app.CommandBars.ExecuteMso("SheetUnhide");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to open the Unhide dialog:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Modal dialog: once closed, sync the pane with the result.
                RefreshSheetList();
            }
        }

        private void SelectAllSheets()
        {
            try
            {
                var wb = ActiveWorkbook();
                if (wb == null) return;

                wb.Sheets.Select(Type.Missing);

                _sheetList.BeginUpdate();
                for (int i = 0; i < _sheetList.Items.Count; i++)
                    _sheetList.SetSelected(i, true);
                _sheetList.EndUpdate();
            }
            catch { }
        }

        private void UngroupSheets()
        {
            try
            {
                var wb = ActiveWorkbook();
                if (wb == null) return;

                if (_sheetList.SelectedIndices.Count > 0)
                {
                    int first = _sheetList.SelectedIndices[0];
                    if (_sheetList.Items[first] is SheetTabItem firstItem)
                    {
                        var sheet = wb.Sheets[firstItem.Name] as Excel.Worksheet;
                        sheet?.Activate();
                    }
                }

                _sheetList.ClearSelected();

                var app = Globals.Application;
                string activeName = null;
                if (app?.ActiveSheet is Excel.Worksheet aws) activeName = aws.Name;
                else if (app?.ActiveSheet is Excel.Chart ach) activeName = ach.Name;

                if (activeName != null)
                {
                    for (int i = 0; i < _sheetList.Items.Count; i++)
                    {
                        if (_sheetList.Items[i] is SheetTabItem item && item.Name == activeName)
                        {
                            _sheetList.SetSelected(i, true);
                            break;
                        }
                    }
                }
            }
            catch { }
        }

        private void DeleteSelectedSheet()
        {
            var wb = ActiveWorkbook();
            if (wb == null) return;

            var toDelete = SelectedSheetNames();

            if (toDelete.Count == 0) return;

            var confirm = MessageBox.Show(
                $"Delete {toDelete.Count} sheet(s)?\n\n" + string.Join(", ", toDelete) +
                "\n\nThis action cannot be undone.",
                "SheetDock", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (confirm != DialogResult.Yes) return;

            try
            {
                foreach (var name in toDelete)
                {
                    object sheet = wb.Sheets[name];

                    if (sheet is Excel.Worksheet ws)
                    {
                        if (wb.Worksheets.Count <= 1)
                        {
                            MessageBox.Show("Cannot delete the only worksheet.",
                                "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        ws.Delete();
                    }
                    else if (sheet is Excel.Chart ch)
                    {
                        if (wb.Sheets.Count <= 1) return;
                        ch.Delete();
                    }
                }

                RefreshSheetList();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to delete:\n" + ex.Message,
                    "SheetDock", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ========================================================
        //  IDisposable
        // ========================================================
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _instances.Remove(this);

            _sheetList.DoubleClick -= SheetList_DoubleClick;
            _sheetList.SelectedIndexChanged -= SheetList_SelectedIndexChanged;
            _sheetList.MouseDown -= SheetList_MouseDown;
            _sheetList.MouseMove -= SheetList_MouseMove;
            _sheetList.MouseUp -= SheetList_MouseUp;
            _sheetList.KeyDown -= SheetList_KeyDown;

            _autoScrollTimer?.Stop();
            _autoScrollTimer?.Dispose();

            _themeTimer?.Stop();
            _themeTimer?.Dispose();

            _contextMenu?.Dispose();
            _host.Dispose();
        }
    }

    // ============================================================
    //  DIALOG FORMS
    // ============================================================
    public class SheetPickerForm : Form
    {
        private readonly CheckedListBox _list;
        public List<string> SelectedSheetNames { get; } = new List<string>();

        public SheetPickerForm(Excel.Workbook srcWb)
        {
            Text = "Select Sheets to Insert";
            Size = new Size(320, 420);
            StartPosition = FormStartPosition.CenterScreen;

            _list = new CheckedListBox
            {
                Dock = DockStyle.Fill,
                CheckOnClick = true,
                Font = new Font("Segoe UI", 9f)
            };

            foreach (object sheetObj in srcWb.Sheets)
            {
                if (sheetObj is Excel.Worksheet ws)
                    _list.Items.Add(ws.Name);
            }

            var btnPanel = new Panel { Dock = DockStyle.Bottom, Height = 40 };
            var ok = new Button { Text = "OK", Dock = DockStyle.Right, Width = 80, DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", Dock = DockStyle.Right, Width = 80, DialogResult = DialogResult.Cancel };

            ok.Click += (s, e) =>
            {
                SelectedSheetNames.Clear();
                foreach (var item in _list.CheckedItems)
                    SelectedSheetNames.Add(item.ToString());
            };

            btnPanel.Controls.Add(ok);
            btnPanel.Controls.Add(cancel);

            Controls.Add(_list);
            Controls.Add(btnPanel);
            AcceptButton = ok;
            CancelButton = cancel;
        }
    }

    // ============================================================
    //  HOST
    // ============================================================
    [ComVisible(true)]
    public class NavigatorPaneHost : UserControl
    {
        private readonly NavigatorPane _pane;

        /// <summary>Raised when the host size changes (e.g. the pane is docked to another side).</summary>
        public static event Action HostResized;

        public NavigatorPaneHost()
        {
            _pane = new NavigatorPane();
            this.Dock = DockStyle.Fill;
            this.Controls.Add(_pane.HostControl);
            this.Resize += (s, e) => HostResized?.Invoke();
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
                _pane?.Dispose();

            base.Dispose(disposing);
        }
    }
}