using System;
using System.Runtime.InteropServices;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Excel = Microsoft.Office.Interop.Excel;

namespace SheetDock
{
    [ComVisible(true)]
    public class SheetDockRibbon : ExcelRibbon
    {
        public override string GetCustomUI(string RibbonID)
        {
            return @"<?xml version=""1.0"" encoding=""UTF-8""?>
<customUI xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">
  <ribbon>
    <tabs>
      <tab id=""SheetDockTab"" label=""SheetDock"">
        <group id=""SheetDockGroup"" label=""View"">
          <button id=""btnTogglePane""
                  label=""Custom SheetDock""
                  size=""large""
                  imageMso=""ViewTwoWindows""
                  onAction=""OnTogglePane""
                  screentip=""Show / hide the custom SheetDock pane""
                  supertip=""Custom sheet navigation pane.""/>
          <button id=""btnToggleTabs""
                  label=""Built-in Excel Tabs""
                  size=""large""
                  imageMso=""TableStyleAddGallery""
                  onAction=""OnToggleTabs""
                  screentip=""Show / hide Excel's built-in sheet tabs""
                  supertip=""Controls the visibility of Excel's built-in worksheet tabs.""/>
          <button id=""btnToggleStats""
                  label=""Quick Stats""
                  size=""large""
                  imageMso=""AutoSum""
                  onAction=""OnToggleStats""
                  screentip=""Show / hide Quick Stats""
                  supertip=""Selection statistics in the status bar; drag a statistic onto a cell to turn it into a formula.""/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        public void OnTogglePane(IRibbonControl control)
        {
            try { CombinedAddIn.TogglePane(); }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error toggle pane:\n" + ex.GetType().Name + "\n" + ex.Message,
                    "SheetDock");
            }
        }

        public void OnToggleTabs(IRibbonControl control)
        {
            try
            {
                var app = Globals.Application;
                if (app?.ActiveWindow == null) return;

                app.ActiveWindow.DisplayWorkbookTabs = !app.ActiveWindow.DisplayWorkbookTabs;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Error toggle tabs:\n" + ex.GetType().Name + "\n" + ex.Message,
                    "SheetDock");
            }
        }

        public void OnToggleStats(IRibbonControl control)
        {
            try
            {
                // StatDock remains a separate module, but its control is unified
                // on the SheetDock ribbon through an Excel-DNA command.
                ((Excel.Application)ExcelDnaUtil.Application).Run("StatDock_Toggle");
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    "Quick Stats is not available.\n" + ex.GetType().Name + "\n" + ex.Message,
                    "SheetDock");
            }
        }
    }
}
