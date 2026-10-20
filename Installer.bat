@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Install SheetDock - Unblock + Stop Excel + Install

:: ============================================================
:: SheetDock / StatDock - Safe Installer / Recovery
:: 1. Request Administrator privilege
:: 2. Stop EXCEL.EXE forcibly
:: 3. Unblock downloaded XLL
:: 4. Install XLL to the user's Excel AddIns folder
:: 5. Do NOT use COM, RegisterXLL, or AddIns.Add
:: ============================================================

:: ---- Self-elevate ------------------------------------------------
net session >nul 2>&1
if not "%errorlevel%"=="0" (
    echo Requesting Administrator privilege...
    powershell -NoProfile -ExecutionPolicy Bypass -Command ^
      "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

echo.
echo ============================================================
echo        SHEETDOCK / STATDOCK - SAFE INSTALLER
echo ============================================================
echo.

set "SRC=%~dp0SheetDock64-packed.xll"
set "DESTDIR=%APPDATA%\Microsoft\AddIns"
set "DEST=%DESTDIR%\SheetDock64-packed.xll"

if not exist "%SRC%" (
    echo ERROR: XLL file not found:
    echo   %SRC%
    echo.
    echo Place this BAT in the same folder as:
    echo   SheetDock64-packed.xll
    echo.
    pause
    exit /b 1
)

:: ---- Stop Excel BEFORE touching the add-in -----------------------
echo [1/4] Stopping Excel...
tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    taskkill /F /IM EXCEL.EXE >nul 2>&1
    timeout /t 3 /nobreak >nul
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul
if not errorlevel 1 (
    echo.
    echo ERROR: EXCEL.EXE is still running after taskkill.
    echo.
    echo Stop Excel / restart Windows, then run this BAT again.
    echo.
    pause
    exit /b 2
)

echo Excel has stopped.

:: ---- Unblock downloaded file ------------------------------------
echo.
echo [2/4] Unblocking XLL file...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unblock-File -LiteralPath '%SRC%' -ErrorAction Stop"
if errorlevel 1 (
    echo.
    echo ERROR: Failed to unblock the XLL.
    echo File:
    echo   %SRC%
    echo.
    pause
    exit /b 3
)
echo XLL unblocked successfully.

:: ---- Install -----------------------------------------------------
echo.
echo [3/4] Installing XLL...
if not exist "%DESTDIR%" mkdir "%DESTDIR%" >nul 2>&1

copy /Y "%SRC%" "%DEST%" >nul
if errorlevel 1 (
    echo.
    echo ERROR: Failed to copy XLL to:
    echo   %DEST%
    echo.
    pause
    exit /b 4
)

:: Also unblock the installed copy.
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Unblock-File -LiteralPath '%DEST%' -ErrorAction Stop"
if errorlevel 1 (
    echo.
    echo WARNING: XLL was copied but unblocking the destination file failed.
    echo   %DEST%
    echo.
)

echo XLL installed at:
echo   %DEST%

:: ---- Final recovery kill, exactly the requested approach --------
echo.
echo [4/4] Recovery: making sure Excel has really stopped...
tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul
if errorlevel 1 (
    echo Excel is not running.
) else (
    echo Closing Excel process...
    taskkill /F /IM EXCEL.EXE >nul 2>&1
    timeout /t 3 /nobreak >nul
)

tasklist /FI "IMAGENAME eq EXCEL.EXE" | find /I "EXCEL.EXE" >nul
if errorlevel 1 (
    echo.
    echo ============================================================
    echo INSTALLATION COMPLETE
    echo ============================================================
    echo.
    echo Excel was stopped successfully.
    echo XLL has been unblocked and installed.
    echo.
    echo IMPORTANT:
    echo - This BAT does NOT use Excel COM.
    echo - This BAT does NOT call RegisterXLL.
    echo - This BAT does NOT call AddIns.Add.
    echo - The XLL is NOT force auto-loaded by this BAT.
    echo.
    echo Open Excel normally.
    echo Enable SheetDock through the Excel Add-in Manager if needed.
    echo.
) else (
    echo.
    echo WARNING: EXCEL.EXE is still running.
    echo Try restarting Windows if the process still cannot be stopped.
    echo.
)

pause
exit /b 0
