@echo off
REM App86 強制全件同期用バッチ
cd /d "%~dp0"

set "POWERSHELL_EXE=C:\Program Files\PowerShell\7\pwsh.exe"
if not exist "%POWERSHELL_EXE%" set "POWERSHELL_EXE=C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe"

"%POWERSHELL_EXE%" -ExecutionPolicy Bypass -NoProfile -File "%~dp0Main.ps1" -ForceFullSync
exit /b %ERRORLEVEL%
