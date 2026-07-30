@echo off
REM Parque Arauco: un Word + hallazgos Q3 por mall.
REM Por defecto SIN enviar correo (RUN automatico local).
REM Para enviar a anibal.aoperaciones@wes.cl: ejecutar el .py sin --no-email o usar run_pa_todos_malls_hallazgos_correo.cmd
cd /d "%~dp0"
python generar_pa_todos_malls_hallazgos_y_enviar.py --no-email %*
set EXITCODE=%ERRORLEVEL%
if not "%EXITCODE%"=="0" echo [ERROR] Codigo salida: %EXITCODE%
exit /b %EXITCODE%
