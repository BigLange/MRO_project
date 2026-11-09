@echo off
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-MRO.ps1"
set "MRO_INSTALL_EXIT=%ERRORLEVEL%"
if not "%MRO_INSTALL_EXIT%"=="0" pause
exit /b %MRO_INSTALL_EXIT%
