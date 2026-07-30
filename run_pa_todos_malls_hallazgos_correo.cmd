@echo off
REM Misma generacion que run_pa_todos_malls_hallazgos.cmd pero CON envio de correo.
cd /d "%~dp0"
python generar_pa_todos_malls_hallazgos_y_enviar.py %*
set EXITCODE=%ERRORLEVEL%
if not "%EXITCODE%"=="0" echo [ERROR] Codigo salida: %EXITCODE%
exit /b %EXITCODE%
