@echo off
rem Tutup Excel, hapus build lama, lalu build ulang (hasil .xll di folder ini).
cd /d "%~dp0"
call npm run rebuild
echo.
pause
