@echo off
cls

set "ORIGEM=\\servidor\compartilhamento\MAPAS PDF"
set "DESTINO=G:\Meu Drive\MAPAS\MAPAS PDF"
set "LOGDIR=G:\Meu Drive\MAPAS\_LOGS"

if not exist "%LOGDIR%" mkdir "%LOGDIR%"
if not exist "%DESTINO%" mkdir "%DESTINO%"

set "DATA=%date:~6,4%-%date:~3,2%-%date:~0,2%"
set "HORA=%time:~0,2%-%time:~3,2%-%time:~6,2%"
set "HORA=%HORA: =0%"

echo ==============================================================
echo                  SINCRONIZACAO DE ARQUIVOS
echo ==============================================================
echo.
echo Origem  : %ORIGEM%
echo Destino : %DESTINO%
echo Log     : %LOGDIR%\Backup_%DATA%_%HORA%.txt
echo.
echo --------------------------------------------------------------
echo Iniciando processo...
echo --------------------------------------------------------------
echo.

robocopy "%ORIGEM%" "%DESTINO%" /MIR /R:2 /W:2 /TEE /LOG:"%LOGDIR%\Backup_%DATA%_%HORA%.txt"

echo.
echo Codigo de retorno do Robocopy: %ERRORLEVEL%
echo --------------------------------------------------------------
echo Processo finalizado.
echo --------------------------------------------------------------
echo.
pause
