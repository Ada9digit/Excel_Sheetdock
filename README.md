# SheetDock + StatDock

**Two small tools that make Excel feel roomier: tidy sheet navigation on the side of your screen, and selection statistics you can simply "drag" into a formula.**

![Platform](https://img.shields.io/badge/platform-Excel%20for%20Windows-217346?logo=microsoftexcel&logoColor=white)
![Framework](https://img.shields.io/badge/.NET%20Framework-4.8-512BD4?logo=dotnet&logoColor=white)
![Excel-DNA](https://img.shields.io/badge/Excel--DNA-1.7.0-blue)
![Build](https://img.shields.io/badge/build-npm%20run%20build-CB3837?logo=npm&logoColor=white)

---

## What is this, and why might you need it?

If you have ever worked in an Excel workbook with **dozens of sheets**, you know these two complaints:

1. **The sheet tabs at the bottom are cramped and awkward.** Names get truncated, you scroll left and right, and finding one sheet among 30 tabs is tiring.
2. **The status bar statistics are view-only.** Excel shows Average, Count, and Sum for your selection, but to actually use those numbers you still have to type the formula yourself.

**SheetDock + StatDock** solves both in a single add-in:

| | The problem in plain Excel | With this add-in |
|---|---|---|
| **SheetDock** | Small, cramped tabs that need scrolling | A side panel listing every sheet as a "chip", reorderable by drag |
| **StatDock** | Statistics are displayed but can't be reused | Statistic chips in the status bar that you can **drag onto a cell** to create a formula |

There is nothing to learn. Install it, open Excel, and everything shows up right away.

---

## A quick look

A rough sketch (not an actual screenshot):

```text
┌─────────────┬──────────────────────────────────────────────────────┐
│ SheetDock   │   A      B       C       D                           │
│             │ 1  Name   Qty    Price   Total                       │
│ ⠿ Summary   │ 2  Coffee 10     15000   150000                      │
│ ⠿ Sales     │ 3  Tea    8      12000    96000                      │
│ ⠿ Stock     │ 4  Milk   5      18000    90000                      │
│ ⠿ Reports   │ 5                                                    │
│ ⠿ Data 2024 │                                                      │
│ ⠿ Data 2025 │                                                      │
│    ...      │                                                      │
├─────────────┴──────────────────────────────────────────────────────┤
│ Ready   Sheet 2 of 12      Average  Count  Min  Max  [ Sum ]  100% │
└────────────────────────────────────────────────────────────────────┘
   ▲ SheetDock: the sheet list         ▲ StatDock: drag a chip onto a cell
```

---

## SheetDock: sane sheet navigation

A panel on the left side of Excel (it can be docked elsewhere) that shows **all your sheets as a list of chips**.

- **One click to switch sheets.** No tab scrolling. Both worksheets and chart sheets are supported.
- **Drag to reorder.** Pull a chip by the handle on its side to move the sheet. The list auto-scrolls when you drag toward an edge.
- **Select several sheets at once** with Ctrl or Shift, for example to duplicate multiple sheets in one go.
- **Tab colors are shown**, so the color-coding you already use on your sheets stays useful.
- **Right-click** opens Excel's native sheet menu (Rename, Delete, Move or Copy, Tab Color, Hide, Unhide, and so on), plus four extra items:
  - **Add Sheet**
  - **Insert From File...** (insert a sheet from another Excel file)
  - **Duplicate**
  - **Duplicate Selection** (duplicate every selected sheet)
- **Keyboard shortcuts:** `F2` to rename, `Delete` to delete the selected sheet.
- **Hidden sheets don't clutter the list.**
- **The look follows your Office theme** (White, Colorful, Dark Gray, or Black), so it doesn't feel like a foreign object.
- **Excel's built-in tabs are hidden** to avoid duplicates. You can bring them back any time from the ribbon.

---

## StatDock: statistics you can drag

A thin strip that sits on the **Excel status bar** (where the built-in Average/Count/Sum normally appear) and shows statistics for your current selection:

`Average` · `Count` · `Numerical Count` · `Min` · `Max` · `Sum`

What makes it different from the regular status bar:

- **Drag a chip onto a cell to get a formula.** Select A1:A10, drag the **Sum** chip onto an empty cell, and that cell contains `=SUM(A1:A10)`. While dragging, a small preview shows the address of the target cell.
- **Filter and hidden-row aware.** The numbers shown come only from **visible** cells. The generated formula adapts: for filtered data it uses the addresses of the visible areas, and if that gets too long it uses `SUBTOTAL(109, ...)`, which skips hidden rows automatically.
- **Guards against mistakes.** Dropping a chip inside its own source range is rejected (avoiding a *circular reference*), and so is dropping it anywhere that isn't a cell.
- **Adapts to the width.** When space is tight, the less important chips are hidden first. `Sum` is never hidden.
- **Lightweight.** Updates run after a short delay (debounce) and are skipped while you drag-select or while Excel is calculating, so Excel doesn't feel sluggish.
- **Follows the theme and status bar.** Colors match Excel, and the strip disappears if the status bar is turned off.

---

## Ribbon

The add-in adds one **SheetDock** tab to the ribbon with three buttons:

| Button | What it does |
|---|---|
| **Custom SheetDock** | Show or hide the sheet list panel |
| **Built-in Excel Tabs** | Show or hide Excel's built-in sheet tabs |
| **Quick Stats** | Turn the StatDock strip in the status bar on or off |

---

## Requirements

- **Windows** with **Microsoft Excel desktop** (not Excel for Mac or Excel in the browser)
- **.NET Framework 4.8** (already included in recent Windows 10 and 11)
- Two file variants are provided: **64-bit** (`SheetDock64-packed.xll`) and **32-bit** (`SheetDock-packed.xll`). Pick the one that matches your Excel. Most modern Excel installations are 64-bit.

---

## Installation

### Quick way (recommended)

1. Make sure `SheetDock64-packed.xll` and `Installer.bat` are in the same folder.
2. **Double-click `Installer.bat`.** Allow the Administrator prompt.
3. The installer closes Excel (save your work first), unblocks the downloaded file, and copies the add-in into your own AddIns folder.
4. Open Excel. If the **SheetDock** tab doesn't appear, enable it via *File → Options → Add-ins → Manage: Excel Add-ins → Go*.

### Manual way

1. Open Excel, then go to *File → Options → Add-ins*.
2. Choose **Manage: Excel Add-ins**, click **Go...**, then **Browse...**.
3. Select the matching `.xll` file and tick the add-in.

> **If Windows blocks the file:** right-click the `.xll` file → *Properties* → tick **Unblock** → OK. `Installer.bat` does this automatically.

---

## Building from source

For developers. All you need is **Windows** and **Node.js 16+**. The `dotnet` command doesn't have to be installed already: if it isn't found, the script downloads a portable .NET SDK into a `.dotnet/` folder inside the project (once only, no admin rights, and it doesn't touch your PATH).

```powershell
npm run build             # Release build; the .xll files appear in the project folder
npm run rebuild           # close Excel, delete the old build, then build again
npm run rebuild:install   # same as rebuild, then run Installer.bat straight away
npm run clean             # delete bin, obj, and the built .xll files
```

No `npm install` is needed: the script uses only Node's built-in modules.

Prefer not to open a terminal? Double-click:

- `rebuild.bat` → close Excel, clean, rebuild
- `rebuild-install.bat` → everything above, then install into Excel

> ⚠️ `rebuild` and `rebuild:install` **force-close Excel** so the old `.xll` isn't locked. Save your work first.

The build produces two files in the project root:

```text
SheetDock64-packed.xll   Excel 64-bit
SheetDock-packed.xll     Excel 32-bit
```

These `.xll` files are *packed* (the DLL and dependencies are embedded), so one file per architecture is enough to share.

---

## Project structure

```text
SheetDock/
├── App/
│   └── CombinedAddIn.cs          Add-in entry point, wires both modules together
├── UI/
│   └── SheetDockRibbon.cs        Ribbon tab and buttons
├── SheetDock/Navigation/         SheetDock module
│   ├── NavigatorPane.cs          The sheet list panel (UI, drag, menu, sheet actions)
│   ├── RecentSheets.cs           History of visited sheets
│   ├── FolderStore.cs            Storage for a virtual folder structure inside the workbook
│   ├── Commands.cs               Macros for the right-click menu
│   └── Globals.cs
├── StatDock/                     StatDock module
│   ├── Core/
│   │   ├── StatEngine.cs         Reads the selection and computes statistics
│   │   ├── StatHost.cs           Lifecycle, debounce, polling, refresh guards
│   │   ├── Commands.cs           Puts the formula into the cell
│   │   └── Globals.cs
│   ├── UI/
│   │   ├── StatBar.cs            The strip on the status bar
│   │   ├── StatChip.cs           The draggable statistic chips
│   │   ├── DragGhost.cs          Preview shown while dragging
│   │   └── ThemeColors.cs        Colors that follow the Excel theme
│   └── Infrastructure/
│       └── Native.cs             Win32 calls (window position, DPI, mouse)
├── scripts/build.js              npm build script
├── Installer.bat                 Add-in installer
├── rebuild.bat / rebuild-install.bat
├── SheetDock.csproj              .NET Framework 4.8 + Excel-DNA project
└── package.json
```

---

## How it works (for the curious)

**Architecture.** The add-in is built with [Excel-DNA](https://excel-dna.net/), which bridges C# to Excel. `CombinedAddIn` is the single entry point and runs SheetDock and StatDock as two modules inside the same add-in.

**StatDock refresh model.** Reading the selection over COM is expensive, so StatDock uses three timers:

1. **Debounce:** merges bursts of Excel events into a single update.
2. **Track:** light and frequent; it only keeps the strip's position and visibility in sync, without reading any data.
3. **Poll:** detects changes that don't always raise an event (such as filtering or hiding rows), and is rate-limited.

Several guards prevent updates from triggering each other in a chain: a *mute* window after each update, deferral while a mouse button is held, and deferral while Excel is busy or calculating.

**Inserting formulas.** The formula goes through the clipboard and a **Ctrl+V sent to Excel**, so Excel handles it as a user action. Excel's copy mode is cancelled first so that what gets pasted is really the formula. A moment later the result is checked; if the cell doesn't contain a valid formula (for example because of localized function names), the formula is written directly through `Range.Formula` using English function names as a fallback. Every drop is logged to `%TEMP%\SheetDock.log`, which is useful for tracing problems.

---

## Troubleshooting

| Symptom | Likely cause and fix |
|---|---|
| The **SheetDock** tab doesn't appear | Enable it via *File → Options → Add-ins → Go*. Make sure the `.xll` variant (32/64-bit) matches your Excel. |
| Excel blocks the add-in | Right-click the `.xll` → Properties → **Unblock**, or run `Installer.bat`. |
| `npm run build` fails because a file is locked | Excel is still holding the `.xll`. Use `npm run rebuild` (it closes Excel automatically). |
| An `EPERM` error appears during clean | Same cause: Excel has a file in `bin` open. Close Excel or run `npm run rebuild`. |
| StatDock chips cover the status bar text | Adjust the `KeepNativeLeftPx` and `KeepNativeRightPx` constants in `StatDock/UI/StatBar.cs`. |
| The StatDock formula doesn't appear | Check `%TEMP%\SheetDock.log`; the last lines explain which path was used and why. |
| Ctrl+C / Ctrl+V do nothing after clicking a sheet in the panel | Make sure you are on the latest version; keyboard focus is returned to Excel after a plain click in the panel. |

---

## Manual test checklist

After building, check at least these:

- [ ] The add-in loads and the **SheetDock** tab appears
- [ ] The sheet list panel is shown and clicking a chip switches sheets
- [ ] Dragging a sheet chip changes the sheet order
- [ ] Right-click shows the Excel menu plus the four extra items
- [ ] The StatDock strip appears on the status bar and its numbers change as the selection changes
- [ ] Apply a filter or hide rows: the StatDock numbers follow
- [ ] Drag the **Sum** chip onto an empty cell: the formula appears and nothing else comes along with it
- [ ] Copy a cell, click another sheet in the panel, then Ctrl+V: still works normally
- [ ] No endlessly spinning loading cursor after a drag

---

## Notes

- For **Excel desktop on Windows** only.
- No license has been chosen yet. Add a `LICENSE` file before sharing this project publicly.
