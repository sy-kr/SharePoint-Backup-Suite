@echo off
REM ---------------------------------------------------------------------------
REM  Run-Backups.cmd - Launcher for the PowerShell backup orchestrator.
REM
REM  This is a thin wrapper so you can run the backup from Task Scheduler,
REM  a shortcut, or by double-clicking.  All configuration lives in the
REM  PowerShell script (Run-Backups.ps1).
REM
REM  -Headless suppresses all console output (recommended for Task Scheduler).
REM  Remove it to see verbose output when running interactively.
REM
REM  Credentials are loaded automatically from Windows Credential Manager.
REM  Run Setup-Credentials.ps1 once (as the same user) to store them.
REM
REM  Edit the paths below if your install location differs.
REM ---------------------------------------------------------------------------

SET PWSH="C:\Program Files\PowerShell\7\pwsh.exe"
SET SCRIPT_DIR=%~dp0
SET SCRIPT=%SCRIPT_DIR%Run-Backups.ps1

%PWSH% -NoProfile -NonInteractive -ExecutionPolicy Bypass -File "%SCRIPT%" -Headless

exit /b %ERRORLEVEL%
