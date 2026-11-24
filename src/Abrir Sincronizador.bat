@echo off
rem Abre a interface do Sincronizador de Mapas PDF
start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0SincronizadorMapas.ps1"
