@echo off
rem Lance l'application Planning CPDL (page web locale + pilotage d'Outlook).
rem La politique d'exécution est contournée pour ce seul lancement.
powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0Planning-Outlook.ps1" %*
if errorlevel 1 pause
