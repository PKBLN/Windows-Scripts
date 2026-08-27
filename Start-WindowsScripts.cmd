@echo off
setlocal

set "WINDOWS_SCRIPTS_BATCH=%~f0"
set "WINDOWS_SCRIPTS_LAUNCHER=%~dp0Start-WindowsScripts.ps1"

if exist "%WINDOWS_SCRIPTS_LAUNCHER%" goto launcher_found
echo Starter nicht gefunden / Launcher not found:
echo "%WINDOWS_SCRIPTS_LAUNCHER%"
pause
exit /b 1

:launcher_found

rem Remove possible Internet zone markers from this batch file and the PowerShell launcher.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command "Unblock-File -LiteralPath $env:WINDOWS_SCRIPTS_BATCH -ErrorAction Stop; Unblock-File -LiteralPath $env:WINDOWS_SCRIPTS_LAUNCHER -ErrorAction Stop"
if errorlevel 1 (
    echo Der Starter konnte nicht freigegeben werden. / The launcher could not be unblocked.
    pause
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%WINDOWS_SCRIPTS_LAUNCHER%"
if errorlevel 1 (
    echo Der Starter wurde mit einem Fehler beendet. / The launcher exited with an error.
    pause
    exit /b 1
)

endlocal
