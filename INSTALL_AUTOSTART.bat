@echo off
rem Installs mail monitor autostart (shortcut in the Startup folder). Run once.
rem ASCII only: all text lives in _launcher.ps1 (cmd.exe misreads UTF-8 .bat files).
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0_launcher.ps1" -Mode autostart-install
pause
