using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatDock
{
    public enum StatKind
    {
        Average = 0,
        Count = 1,
        NumCount = 2,
        Min = 3,
        Max = 4,
        Sum = 5
    }

    public static class StatInfo
    {
        // Same order as Excel's status bar.
        public static readonly StatKind[] Order =
        {
            StatKind.Average, StatKind.Count, StatKind.NumCount,
            StatKind.Min, StatKind.Max, StatKind.Sum
        };

        public static string Label(StatKind k)
        {
            switch (k)
            {
                case StatKind.Average: return "Average";
                case StatKind.Count: return "Count";
                case StatKind.NumCount: return "Numerical Count";
                case StatKind.Min: return "Min";
                case StatKind.Max: return "Max";
                default: return "Sum";
            }
        }

        /// <summary>Function name (English, used for Range.Formula).</summary>
        public static string Function(StatKind k)
        {
            switch (k)
            {
                case StatKind.Average: return "AVERAGE";
                case StatKind.Count: return "COUNTA";     // "Count" in the status bar = non-empty cells
                case StatKind.NumCount: return "COUNT";   // "Numerical Count" = cells containing numbers
                case StatKind.Min: return "MIN";
                case StatKind.Max: return "MAX";
                default: return "SUM";
            }
        }

        /// <summary>SUBTOTAL function_num code that ignores hidden rows.</summary>
        public static int SubtotalCode(StatKind k)
        {
            switch (k)
            {
                case StatKind.Average: return 101;
                case StatKind.Count: return 103;
                case StatKind.NumCount: return 102;
                case StatKind.Min: return 105;
                case StatKind.Max: return 104;
                default: return 109;
            }
        }
    }

    /// <summary>Result of reading the Excel selection at a single point in time.</summary>
    public sealed class SelectionSnapshot
    {
        public bool HasRange;
        public string SelectionAddress = "";
        public double SelectedCells;
        public double VisibleCells;

        /// <summary>The selection contains hidden cells (filtered out / hidden rows or columns).</summary>
        public bool Filtered { get { return VisibleCells < SelectedCells; } }

        /// <summary>References inserted into the formula: one address per area, e.g. "A1", "A5:A15".</summary>
        public List<string> Refs = new List<string>();

        /// <summary>The formula uses SUBTOTAL(10x, ...) because the cell list is too long.</summary>
        public bool UseSubtotal;

        public bool CanBuildFormula;

        /// <summary>Selection fingerprint (address + height + width). Changes when rows are filtered / hidden.</summary>
        public string Signature = "";
        public string ListSeparator = ",";

        /// <summary>Indexed by (int)StatKind. null = no value.</summary>
        public double?[] Values = new double?[6];

        /// <summary>
        /// Builds the formula text. Use separator "," for Range.Formula (always English format),
        /// or the regional separator for display to the user.
        /// </summary>
        public string BuildFormula(StatKind kind, string separator)
        {
            if (UseSubtotal)
            {
                return "=SUBTOTAL(" + StatInfo.SubtotalCode(kind) + separator +
                       string.Join(separator, Refs) + ")";
            }

            return "=" + StatInfo.Function(kind) + "(" + string.Join(separator, Refs) + ")";
        }
    }

    public static class StatEngine
    {
        private const int MaxAreas = 255;            // Excel function argument limit
        private const int MaxFormulaLength = 8000;   // Excel formula length limit = 8192

        /// <summary>
        /// Read the active selection and compute its statistics.
        /// Returns null if Excel is busy / in edit mode (keep the previous display).
        /// All temporary COM objects are released here, because this function is called repeatedly.
        /// </summary>
        public static SelectionSnapshot Capture(Excel.Application app)
        {
            return Capture(app, true);
        }

        /// <summary>
        /// needRefs = false skips building the area list (dozens of COM calls when data is filtered).
        /// False is enough for display; drag uses true because it needs the exact addresses.
        /// </summary>
        public static SelectionSnapshot Capture(Excel.Application app, bool needRefs)
        {
            Excel.Range sel = null;
            Excel.Range vis = null;

            try
            {
                try { sel = app.Selection as Excel.Range; }
                catch { return null; }

                var snap = new SelectionSnapshot();
                if (sel == null) return snap;   // selection is not a range (chart / shape)

                try
                {
                    snap.HasRange = true;
                    snap.ListSeparator = ReadListSeparator(app);
                    snap.SelectionAddress = AddressOf(sel);
                    snap.SelectedCells = CountOf(sel);
                    snap.Signature = SizeKey(snap.SelectionAddress, sel);

                    // Visible cells only (accounting for filters and hidden rows/columns).
                    // SpecialCells on a single cell expands to the whole sheet, so it is skipped.
                    if (snap.SelectedCells > 1)
                    {
                        try { vis = sel.SpecialCells(Excel.XlCellType.xlCellTypeVisible, Type.Missing); }
                        catch (COMException) { vis = null; }
                    }

                    double visible = vis != null ? CountOf(vis) : snap.SelectedCells;
                    if (visible <= 0 || visible > snap.SelectedCells) visible = snap.SelectedCells;
                    snap.VisibleCells = visible;

                    bool filtered = vis != null && visible < snap.SelectedCells;
                    Excel.Range source = filtered ? vis : sel;

                    if (needRefs) BuildRefs(snap, sel, source);
                    else snap.CanBuildFormula = true;   // confirmed when the drag starts
                    FillStats(app, snap, source);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("StatEngine.Capture: " + ex.Message);
                    return null;
                }

                return snap;
            }
            finally
            {
                Release(vis);
                Release(sel);
            }
        }

        private static string SizeKey(string address, Excel.Range r)
        {
            try
            {
                // Range height / width = 0 for hidden rows / columns, so it changes when a filter is applied.
                return address + "|" +
                       Convert.ToString(r.Height, CultureInfo.InvariantCulture) + "|" +
                       Convert.ToString(r.Width, CultureInfo.InvariantCulture);
            }
            catch { return address; }
        }

        /// <summary>
        /// Fingerprint of the current selection at very low cost (no SpecialCells / formulas).
        /// null = Excel is busy / in edit mode.
        /// </summary>
        public static string Signature(Excel.Application app)
        {
            Excel.Range sel = null;
            try
            {
                sel = app.Selection as Excel.Range;
                if (sel == null) return "";
                return SizeKey(AddressOf(sel), sel);
            }
            catch { return null; }
            finally { Release(sel); }
        }

        private static void BuildRefs(SelectionSnapshot snap, Excel.Range sel, Excel.Range source)
        {
            List<string> refs = AreaAddresses(source, MaxAreas);

            int length = 20;
            foreach (var r in refs) length += r.Length + 1;

            if (refs.Count > 0 && length <= MaxFormulaLength && refs.All(r => r.Length > 0))
            {
                snap.Refs = refs;
                snap.UseSubtotal = false;
                snap.CanBuildFormula = true;
                return;
            }

            // Too many areas / too long: use SUBTOTAL(10x, original selection),
            // which automatically skips filtered rows. SUBTOTAL's first argument takes up one slot.
            List<string> selRefs = AreaAddresses(sel, MaxAreas - 1);

            snap.Refs = selRefs;
            snap.UseSubtotal = true;
            snap.CanBuildFormula = selRefs.Count > 0 && selRefs.All(r => r.Length > 0);
        }

        /// <summary>Address of each area. Empty list if the area count is 0 or exceeds the max.</summary>
        private static List<string> AreaAddresses(Excel.Range range, int max)
        {
            var list = new List<string>();
            Excel.Areas areas = null;

            try
            {
                areas = range.Areas;
                int n = areas.Count;
                if (n <= 0 || n > max) return list;

                for (int i = 1; i <= n; i++)
                {
                    Excel.Range a = null;
                    try
                    {
                        a = areas[i];
                        list.Add(AddressOf(a));
                    }
                    finally { Release(a); }
                }
            }
            catch
            {
                list.Clear();
            }
            finally
            {
                Release(areas);
            }

            return list;
        }

        private static void FillStats(Excel.Application app, SelectionSnapshot snap, Excel.Range values)
        {
            Excel.WorksheetFunction wf = app.WorksheetFunction;

            try
            {
                double? count = Try(() => wf.CountA(values));
                double? numCount = Try(() => wf.Count(values));

                snap.Values[(int)StatKind.Count] = count;
                snap.Values[(int)StatKind.NumCount] = numCount;

                // Like the status bar: if there are no numbers, Sum/Average/Min/Max are not shown.
                if (numCount.HasValue && numCount.Value > 0)
                {
                    snap.Values[(int)StatKind.Sum] = Try(() => wf.Sum(values));
                    snap.Values[(int)StatKind.Average] = Try(() => wf.Average(values));
                    snap.Values[(int)StatKind.Min] = Try(() => wf.Min(values));
                    snap.Values[(int)StatKind.Max] = Try(() => wf.Max(values));
                }
            }
            finally
            {
                Release(wf);
            }
        }

        private static double? Try(Func<double> f)
        {
            try { return f(); }
            catch { return null; }
        }

        // ---------- Range helpers ----------

        /// <summary>Release a single COM reference. Safe to call with null / non-COM objects.</summary>
        public static void Release(object o)
        {
            if (o == null) return;
            try
            {
                if (Marshal.IsComObject(o)) Marshal.ReleaseComObject(o);
            }
            catch { }
        }

        public static string AddressOf(Excel.Range r)
        {
            try { return r.Address[false, false]; }
            catch { return ""; }
        }

        private static double CountOf(Excel.Range r)
        {
            try { return Convert.ToDouble((object)r.CountLarge); }
            catch
            {
                try { return r.Count; }
                catch { return 0; }
            }
        }

        private static string ReadListSeparator(Excel.Application app)
        {
            try
            {
                var s = app.International[Excel.XlApplicationInternational.xlListSeparator] as string;
                if (!string.IsNullOrEmpty(s)) return s;
            }
            catch { }
            return ",";
        }

        // ---------- drop target ----------

        /// <summary>
        /// The worksheet cell under a given screen point (pixel), or null.
        /// The caller is responsible for calling Release() on the result.
        /// </summary>
        public static Excel.Range RangeAt(Excel.Application app, Point screenPoint)
        {
            Excel.Window win = null;
            try
            {
                win = app.ActiveWindow;
                if (win == null) return null;

                object o = win.RangeFromPoint(screenPoint.X, screenPoint.Y);
                return o as Excel.Range;
            }
            catch
            {
                return null;
            }
            finally
            {
                Release(win);
            }
        }

        /// <summary>True if the target cell is inside the source selection (would become a circular reference).</summary>
        public static bool IsCircular(Excel.Application app, Excel.Range target)
        {
            Excel.Range sel = null;
            Excel.Range hit = null;

            try
            {
                sel = app.Selection as Excel.Range;
                if (sel == null) return false;

                hit = app.Intersect(target, sel);
                return hit != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                Release(hit);
                Release(sel);
            }
        }
    }
}
