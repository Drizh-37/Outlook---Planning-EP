@echo off
rem Lance l'application Planning CPDL dans sa propre fenêtre, sans console visible.
rem La politique d'exécution est contournée pour ce seul lancement.
rem Diagnostic : "Lancer-Planning-Outlook.cmd /console" garde la console PowerShell affichée.

if /i "%~1"=="/console" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Planning-Outlook.ps1"
    pause
    exit /b
)

rem conhost --headless exécute PowerShell sans aucune fenêtre (y compris avec Windows Terminal)
start "" conhost.exe --headless powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Planning-Outlook.ps1" %*
