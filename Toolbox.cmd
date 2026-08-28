@echo off
rem Double-click entry point for Toolbox.ps1.
rem -STA is required: WPF will not start on a multi-threaded apartment.
rem -ExecutionPolicy Bypass applies to this one process only and changes
rem nothing on the machine, which is the same trick the README documents.
powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%~dp0Toolbox.ps1" %*
if errorlevel 1 (
  echo.
  echo The toolbox closed with an error. The message above says why.
  pause
)
