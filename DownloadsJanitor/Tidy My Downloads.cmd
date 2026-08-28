@echo off
setlocal EnableExtensions DisableDelayedExpansion
title Tidy My Downloads
if not exist "%~dp0tool\DownloadsJanitor.Setup.ps1" goto nofiles
powershell.exe -NoProfile -NoLogo -ExecutionPolicy Bypass -File "%~dp0tool\DownloadsJanitor.Setup.ps1"
goto :eof

:nofiles
echo.
echo   The tool files are missing.
echo.
echo   This usually happens when the folder is opened straight from
echo   inside the zip file. Windows will not run it from there.
echo.
echo   Please do this instead:
echo     1. Right-click the zip file, choose Properties,
echo        tick Unblock near the bottom, then click OK
echo     2. Right-click the zip file again and choose Extract All
echo     3. Open the extracted folder and double-click this file again
echo.
pause
