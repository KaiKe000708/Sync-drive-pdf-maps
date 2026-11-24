@echo off
rem Sincroniza sem abrir janela. Use este arquivo no Agendador de Tarefas.
rem Codigos de saida: 0 = ok, 1 = concluido com falhas, 2 = nao executou
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0SincronizadorMapas.ps1" -Auto
exit /b %ERRORLEVEL%
