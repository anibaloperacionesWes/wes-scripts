@echo off
REM Parque Arauco: Mall El Bosque — Word agregado + hallazgos Q3 (sin correo por defecto).
REM Quitar --no-email para enviar correo.
cd /d "%~dp0"
python generar_pa_todos_malls_hallazgos_y_enviar.py --solo-mall "El Bosque" --no-email %*
set EXITCODE=%ERRORLEVEL%
if not "%EXITCODE%"=="0" echo [ERROR] Codigo salida: %EXITCODE%
exit /b %EXITCODE%
