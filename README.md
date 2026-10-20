# SheetDock + StatDock

An Excel-DNA add-in that combines **SheetDock** and **StatDock** into a single Excel add-in.

## Features

### SheetDock
- Worksheet/workbook navigation from a dock pane.
- Worksheet buttons are displayed as chips.
- Quick navigation between sheets.
- The UI uses a visual theme consistent with StatDock.

### StatDock
- Shows statistics for the current Excel selection in the status bar.
- Event-based refresh with debounce.
- Lightweight polling for changes such as filtering/hiding rows, which don't always raise an Excel event.
- The status bar follows Excel's `DisplayStatusBar`.
- Supports suspending visuals during startup or certain operations.
- Refresh is held back during mouse operations to prevent repeated COM reads.
- Protection against Excel being busy/calculating.
- Native Paste is used to insert formulas so that undo stays managed by Excel.

## Architecture

Main structure:

```text
SheetDock/
├── CombinedAddIn.cs
├── Ribbon/
│   └── ...
├── SheetDock/
│   └── Navigation/
│       └── NavigatorPane.cs
└── StatDock/
    ├── Core/
    │   ├── Commands.cs
    │   ├── StatEngine.cs
    │   └── StatHost.cs
    └── UI/
        └── StatBar.cs
```

`CombinedAddIn` is the add-in entry point. SheetDock and StatDock run as modules inside the same add-in.

## StatDock Refresh Model

StatDock uses three timers:

1. **Debounce timer**
   - Merges consecutive Excel events.
   - Prevents every event from immediately running `StatEngine.Capture()`.

2. **Track timer**
   - Runs more frequently.
   - Used to keep the strip's position/visibility on the status bar in sync.
   - Does not perform a full COM statistics read.

3. **Poll timer**
   - Used to detect changes that don't always raise an Excel event.
   - Examples: changes caused by filtering or hidden rows.
   - Polling is throttled so it doesn't keep calling COM while Excel is busy.

### Refresh Gate

`StatHost` has several protection mechanisms:

- `_updateRunning`
- `_eventMuteUntil`
- `_filterSettleUntil`
- `_nextPollAt`
- `_busyRetryAt`
- `_nextUpdateAt`

The goal is to make sure a single Excel operation doesn't produce a chain like:

```text
Excel event
    ↓
StatDock Update
    ↓
StatEngine.Capture
    ↓
Excel event
    ↓
StatDock Update
    ↓
...
```

Events that occur during the mute/settle period are not used to trigger nested refreshes.

## Native Paste and Undo

StatDock formulas are not written using:

```csharp
target.Formula = formula;
```

Instead, the clipboard + Excel's native command is used:

```text
Clipboard
   ↓
target.Select()
   ↓
CommandBars.ExecuteMso("Paste")
   ↓
Excel native undo history
```

The implementation lives in:

```text
StatDock/Core/Commands.cs
```

This path is used deliberately because writing directly to `Range.Formula` through automation can interfere with Excel's native undo history.

StatDock does not install:

```text
Application.OnUndo
Application.OnRepeat
```

This way, undo/redo commands remain under Excel's native mechanism.

## External Edit Protection

Before the native Paste:

```csharp
StatHost.BeginExternalEdit();
```

StatDock then holds back refreshes while Excel performs the Paste operation.

After the operation completes:

```csharp
StatHost.EndExternalEdit(250);
```

This is necessary because native Paste can raise several Excel events, for example:

- Activate
- Select
- Change
- Calculate

Those events must not immediately trigger nested `StatEngine.Capture()` calls.

## Build (npm)

Prasyarat: Windows dan [Node.js](https://nodejs.org) >= 16. .NET SDK tidak perlu diinstall: jika `dotnet` tidak ada di PATH, `npm run build` mengunduh SDK portable ke folder `.dotnet\\` di project (sekali saja, butuh internet, tanpa admin). Target project adalah .NET Framework 4.8; kompilasi C# tetap dilakukan `dotnet`, sedangkan npm hanya menyiapkan SDK, menjalankan build, dan merapikan hasilnya. Folder `.dotnet` tidak dihapus oleh `npm run clean`; hapus manual jika ingin membersihkannya.

```powershell
npm run build      # build Release, hasil .xll di folder ini
npm run rebuild    # tutup Excel, hapus bin/obj/xll lama, lalu build
npm run rebuild:install  # seperti rebuild, lalu jalankan Installer.bat
npm run clean      # hapus bin, obj, dan *-packed.xll di root
```

Tidak ada `npm install` yang perlu dijalankan karena script tidak memakai dependency.

`npm run rebuild` menutup Excel secara paksa (taskkill), jadi simpan pekerjaanmu dulu. Dua file `.bat` di root bisa di-double-click: `rebuild.bat` dan `rebuild-install.bat`.

## Release

Hasil build disalin langsung ke folder project (root), bukan ke dalam `bin\Release\net48\publish`:

```text
SheetDock\
├── SheetDock64-packed.xll   (Excel 64-bit)
├── SheetDock-packed.xll     (Excel 32-bit)
├── Installer.bat
├── package.json
├── SheetDock.csproj
└── ...
```

File `.xll` sudah *packed* (DLL dan dependency tertanam), jadi cukup satu file per arsitektur. Untuk memasang, jalankan `Installer.bat`, atau load `.xll` lewat Excel Add-ins.

## Installation

1. Build the project in `Release` configuration.
2. Make sure the output is available in the project folder.
3. Open Excel.
4. Load the `.xll` file from the project folder via:
   - `File`
   - `Options`
   - `Add-ins`
   - `Excel Add-ins`
   - `Browse`
5. Select the `.xll` file.
6. Make sure the SheetDock/StatDock ribbon appears.

## Troubleshooting

### Excel keeps showing the loading cursor after drag/paste

Make sure the latest version of `StatHost.cs` is used.

Native Paste does produce several Excel events. `BeginExternalEdit()` and `EndExternalEdit()` must still be called from `Commands.PasteFormulaNative()`.

### Build fails because the `.xll` file is in use

Close Excel.

If it is still locked:

```powershell
Stop-Process -Name EXCEL -Force
```

Then:

```powershell
Remove-Item -Recurse -Force .\bin, .\obj -ErrorAction SilentlyContinue
dotnet build -c Release
```

### `CS0656 Missing compiler required member Microsoft.CSharp.RuntimeBinder...`

Do not use `dynamic` in `Commands.cs` for `CommandBars.ExecuteMso`.

The current implementation uses reflection:

```csharp
MethodInfo
```

so it doesn't require the C# runtime binder.

### Duplicate class

Make sure there is only one implementation of each class.

Structure in use:

```text
StatDock/Core/Commands.cs
StatDock/Core/StatHost.cs
StatDock/UI/StatBar.cs
SheetDock/Navigation/NavigatorPane.cs
```

Avoid leftover old files with the same name/class in other folders.

## Development Notes

### Do not use `Application.OnUndo`

`Application.OnUndo` replaces Excel's undo command with an add-in callback. For a drag/paste workflow that must stay compatible with Excel's native undo/redo, use Excel's native Paste.

### Do not perform COM capture from the Track timer

`Track()` is designed for lightweight, P/Invoke-only operations. The full statistics read is done by `Update()` via `StatEngine.Capture()`.

### Do not remove the mute window without testing

Excel events can fire back-to-back during:

- Paste
- Select
- Change
- Calculate
- Activate
- Filter

Without debounce/mute, these operations can cause repeated refreshes and make Excel look like it is constantly loading.

## Testing Checklist

After building, test at minimum:

- [ ] The add-in loads in Excel.
- [ ] The ribbon appears.
- [ ] SheetDock can switch worksheets.
- [ ] StatDock appears on the status bar.
- [ ] Selection changes → statistics update.
- [ ] Filter a worksheet → statistics update.
- [ ] Hide/unhide rows → statistics update.
- [ ] Drag a formula from StatDock onto a cell.
- [ ] No loading loop after the drag.
- [ ] `Ctrl+Z` undoes the native Paste.
- [ ] `Ctrl+Y` can redo when Excel has a native redo state.
- [ ] Excel doesn't refresh repeatedly when Calculate/Change events occur.
- [ ] Release build succeeds.
- [ ] The `*-packed.xll` files can be used to load the add-in.

## Key Files

| File | Purpose |
|---|---|
| `CombinedAddIn.cs` | Add-in entry point |
| `Commands.cs` | StatDock commands and native formula Paste |
| `StatHost.cs` | Lifecycle, refresh, debounce, polling, and protection |
| `StatEngine.cs` | Capture and statistics calculation |
| `StatBar.cs` | StatDock status bar UI |
| `NavigatorPane.cs` | SheetDock navigation UI |
| `*-packed.xll` | Deployment output |

## Status

The project currently uses a unified add-in architecture:

```text
Excel
  │
  └── CombinedAddIn
        ├── SheetDock
        │     └── NavigatorPane
        │
        └── StatDock
              ├── StatHost
              ├── StatEngine
              ├── StatBar
              └── Commands
```

The main focus of the StatDock implementation is keeping the UI responsive, avoiding nested refreshes when Excel sends back-to-back events, and preserving Excel's native undo/redo for Paste operations.
