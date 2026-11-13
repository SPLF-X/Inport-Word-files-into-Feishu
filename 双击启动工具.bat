@echo off
setlocal
chcp 65001 >nul
pushd "%~dp0"
if errorlevel 1 goto bad_directory
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run_tool.ps1"
set "result=%ERRORLEVEL%"
popd
if not "%result%"=="0" echo [ERROR] Tool launch. Read messages above.
echo.
pause
exit /b %result%
:bad_directory
echo [ERROR] Could not enter the tool directory.
pause
exit /b 1
