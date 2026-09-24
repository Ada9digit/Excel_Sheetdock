using System;
using System.Windows.Forms;
using System.Reflection;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatDock
{
    /// <summary>Excel commands used by StatDock. Formula insertion is performed through
    /// Excel's native Paste command so Excel owns the resulting undo record. No OnUndo/OnRepeat
    /// callbacks are installed.
    /// </summary>
    public static class Commands
    {
        private static MethodInfo _executeMso;
        [ExcelCommand(Name = "StatDock_Toggle")]
        public static void Toggle() => StatHost.Toggle();

        /// <summary>
        /// Inserts a formula through Excel's native Paste command. This is deliberately different
        /// from Range.Formula = value: the latter is an automation edit and clears Excel's native
        /// undo history. Paste is an actual Excel UI command and therefore remains undoable by Excel.
        /// </summary>
        public static bool PasteFormulaNative(Excel.Application app, Excel.Range target, string formula)
        {
            if (app == null || target == null || string.IsNullOrEmpty(formula)) return false;

            IDataObject previousClipboard = null;
            Excel.Worksheet ws = null;
            Excel.Workbook wb = null;

            try
            {
                // Native Paste generates several Excel events (Activate/Select/Change/Calculate).
                // Mute StatDock before touching Excel so those events cannot start a refresh loop.
                StatHost.BeginExternalEdit();

                previousClipboard = Clipboard.GetDataObject();

                ws = target.Worksheet;
                wb = ws != null ? ws.Parent as Excel.Workbook : null;
                if (ws == null || wb == null) return false;

                wb.Activate();
                ws.Activate();
                target.Select();

                // Clipboard text beginning with '=' is interpreted by Excel as a formula.
                Clipboard.SetText(formula, TextDataFormat.UnicodeText);

                object commandBars = app.CommandBars;
                if (commandBars == null) return false;

                if (_executeMso == null || _executeMso.DeclaringType != commandBars.GetType())
                    _executeMso = commandBars.GetType().GetMethod(
                        "ExecuteMso", new[] { typeof(string) });

                if (_executeMso == null) return false;

                _executeMso.Invoke(commandBars, new object[] { "Paste" });
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("StatDock native paste: " + ex.Message);
                return false;
            }
            finally
            {
                try
                {
                    if (previousClipboard != null)
                        Clipboard.SetDataObject(previousClipboard, true);
                }
                catch { }

                StatEngine.Release(wb);
                StatEngine.Release(ws);

                // Start the settle window only after our own cleanup is complete.
                // Excel's native Paste remains the owner of the undo record.
                StatHost.EndExternalEdit(250);
            }
        }
    }
}
