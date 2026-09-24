using ExcelDna.Integration;

namespace SheetDock
{
    /// <summary>
    /// Macros invoked by the extra items in the native "Ply" context menu
    /// (Excel's built-in sheet tab right-click menu). Those items are installed temporarily
    /// only while the menu is showing; see NavigatorPane.ShowNativeSheetMenu.
    /// </summary>
    public static class PanelCommands
    {
        [ExcelCommand(Name = "SheetDock_AddSheet")]
        public static void AddSheet() => NavigatorPane.ContextPane?.CmdAddSheet();

        [ExcelCommand(Name = "SheetDock_InsertFromFile")]
        public static void InsertFromFile() => NavigatorPane.ContextPane?.CmdInsertFromFile();

        [ExcelCommand(Name = "SheetDock_Duplicate")]
        public static void Duplicate() => NavigatorPane.ContextPane?.CmdDuplicate();

        [ExcelCommand(Name = "SheetDock_DuplicateSelection")]
        public static void DuplicateSelection() => NavigatorPane.ContextPane?.CmdDuplicateSelection();
    }
}
