@echo off
pushd "%~dp0"
powershell -ExecutionPolicy Bypass -File "%~dp0Install-MRO.ps1"
set EXIT_CODE=%ERRORLEVEL%
popd
if not "%EXIT_CODE%"=="0" (
    echo.
    echo Install failed. Please extract the zip package first, then run this file again.
    pause
    exit /b %EXIT_CODE%
)
echo.
echo Install finished.
pause
