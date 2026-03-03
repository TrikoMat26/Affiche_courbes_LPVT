@echo off
:: Change le dossier courant au dossier du script
cd /d "%~dp0"

:: Nom du script PowerShell
set SCRIPT_NAME=AfficheCourbeVPRT.ps1

:: Vérifie que le script PowerShell existe
if not exist "%SCRIPT_NAME%" (
    echo ❌ Le fichier %SCRIPT_NAME% est introuvable.
    pause
    exit /b 1
)

:: Lance le script PowerShell avec exécution autorisée temporairement
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_NAME%"

pause
