using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using Excel = Microsoft.Office.Interop.Excel;

namespace StatDock
{
    /// <summary>Excel commands used by StatDock. The formula is inserted by sending a real Ctrl+V
    /// keystroke to Excel, so the paste is a user-input operation and stays in Excel's native Undo
    /// history (ExecuteMso / Range.Formula run in automation context and clear the Undo stack).
    /// No OnUndo/OnRepeat callbacks are installed.
    /// </summary>
    public static class Commands
    {
        [ExcelCommand(Name = "StatDock_Toggle")]
        public static void Toggle() => StatHost.Toggle();

        /// <summary>Which path the last paste took (shown in the StatDock flash and logged).</summary>
        public static string LastNote = "";

        private static void Note(string text)
        {
            LastNote = text;
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "SheetDock.log"),
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  " + text + Environment.NewLine);
            }
            catch { }
        }

        private const int ErrNameCode = -2146826259;   // #NAME?

        // State for the delayed check that runs after Excel has processed the Ctrl+V.
        private static Timer _verifyTimer;
        private static Excel.Worksheet _verifyWs;
        private static string _verifyAddress;
        private static string _verifyEnglish;

        private static bool HasFormula(Excel.Range r)
        {
            try { object v = r.HasFormula; return v is bool && (bool)v; }
            catch { return false; }
        }

        /// <summary>
        /// Puts the formula text on the clipboard and sends Ctrl+V to Excel. Excel processes the
        /// keystroke after this call returns; a short timer then checks the result and, only if the
        /// paste did not produce a working formula, writes englishFormula through Range.Formula
        /// (English names, "," separator; this fallback is not undoable).
        /// </summary>
        public static bool PasteFormulaNative(Excel.Application app, Excel.Range target, string formula, string englishFormula = null)
        {
            if (app == null || target == null || string.IsNullOrEmpty(formula)) return false;

            LastNote = "";
            Excel.Worksheet ws = null;
            Excel.Workbook wb = null;
            bool keepWs = false;

            try
            {
                // Paste raises several Excel events (Activate/Select/Change/Calculate).
                // Mute StatDock before touching Excel so those events cannot start a refresh loop.
                StatHost.BeginExternalEdit();

                ws = target.Worksheet;
                wb = ws != null ? ws.Parent as Excel.Workbook : null;
                if (ws == null || wb == null) return false;

                string address = target.Address[false, false];

                wb.Activate();
                ws.Activate();
                target.Select();

                // While Excel is in copy mode (marching ants) Paste uses the copied range, not the
                // clipboard text, which made other cells "follow" the formula. Cancel copy mode first.
                try { app.CutCopyMode = (Excel.XlCutCopyMode)0; } catch { }

                // Text beginning with '=' is interpreted by Excel as a formula.
                Clipboard.SetDataObject(formula, true, 10, 50);

                app.SendKeys("^v", false);

                // Verify after Excel has handled the keystroke.
                StopVerify();
                _verifyWs = ws;
                keepWs = true;
                _verifyAddress = address;
                _verifyEnglish = englishFormula;
                _verifyTimer = new Timer { Interval = 450 };
                _verifyTimer.Tick += VerifyTick;
                _verifyTimer.Start();

                Note("Ctrl+V sent to " + address);
                return true;
            }
            catch (Exception ex)
            {
                Note("paste error: " + ex.GetType().Name + " " + ex.Message.Trim());
                return false;
            }
            finally
            {
                StatEngine.Release(wb);
                if (!keepWs) StatEngine.Release(ws);

                // The keystroke is handled after we return, so keep the mute window longer.
                StatHost.EndExternalEdit(700);
            }
        }

        private static void StopVerify()
        {
            if (_verifyTimer != null)
            {
                try { _verifyTimer.Stop(); _verifyTimer.Dispose(); } catch { }
                _verifyTimer = null;
            }
            if (_verifyWs != null)
            {
                StatEngine.Release(_verifyWs);
                _verifyWs = null;
            }
        }

        private static void VerifyTick(object sender, EventArgs e)
        {
            try { _verifyTimer.Stop(); } catch { }

            Excel.Range r = null;
            try
            {
                r = _verifyWs.Range[_verifyAddress];

                bool hasFormula = HasFormula(r);
                object v = null;
                try { v = r.Value2; } catch { }
                bool nameError = v is int && (int)v == ErrNameCode;

                if ((!hasFormula || nameError) && !string.IsNullOrEmpty(_verifyEnglish))
                {
                    StatHost.BeginExternalEdit();
                    try { r.Formula = _verifyEnglish; }
                    finally { StatHost.EndExternalEdit(250); }
                    Note("FALLBACK Range.Formula at " + _verifyAddress + " (no undo). Reason: " +
                         (nameError ? "#NAME? after paste (localized function names)" : "paste did not produce a formula"));
                }
                else
                {
                    Note("Ctrl+V paste OK at " + _verifyAddress);
                }
            }
            catch (Exception ex)
            {
                Note("verify error: " + ex.GetType().Name + " " + ex.Message.Trim());
            }
            finally
            {
                StatEngine.Release(r);
                StopVerify();
            }
        }
    }
}
