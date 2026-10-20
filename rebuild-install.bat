@echo off
rem Tutup Excel, hapus build lama, build ulang, lalu jalankan Installer.bat.
cd /d "%~dp0"
call npm run rebuild:install
echo.
pause
