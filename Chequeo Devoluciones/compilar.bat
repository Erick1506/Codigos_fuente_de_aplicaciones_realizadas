@echo off
REM compilar.bat - Lanza "streamlit run lista.py" con puerto libre (8501-8510)
pushd "%~dp0"
setlocal enabledelayedexpansion

REM Si tienes streamlit instalado globalmente (en PATH), no hace falta python -m
set "STREAMLIT_EXE=streamlit"

set "SCRIPT=%~dp0lista.py"
set "LOGFILE=%~dp0streamlit_run.log"
set "LAUNCHER=%~dp0_launch_streamlit.cmd"

if not exist "%SCRIPT%" (
  echo ERROR: no se encontro el archivo: %SCRIPT%
  pause
  popd
  exit /b 1
)

if exist "%LOGFILE%" del /f /q "%LOGFILE%" >nul 2>&1

REM Buscar puerto libre entre 8501 y 8510
set "PORT="
for /L %%P in (8501,1,8510) do (
  netstat -ano ^| findstr ":%%P " ^| findstr LISTENING >nul 2>&1
  if errorlevel 1 (
    set "PORT=%%P"
    goto :FOUNDPORT
  )
)
echo No se encontro puerto libre entre 8501 y 8510.
pause
popd
exit /b 2

:FOUNDPORT
echo Puerto elegido: %PORT%

REM Crear lanzador temporal con el comando streamlit directo
(
  echo @echo off
  echo echo Iniciando Streamlit en puerto %PORT%...
  echo echo Comando: %STREAMLIT_EXE% run "%SCRIPT%" --server.port %PORT%
  echo.
  echo "%STREAMLIT_EXE%" run "%SCRIPT%" --server.port %PORT% ^>^> "%LOGFILE%" 2^>^&1
  echo.
  echo echo.
  echo echo Logs guardados en: "%LOGFILE%"
  echo echo Presiona una tecla para cerrar esta ventana cuando quieras detener.
  echo pause
) > "%LAUNCHER%"

REM Ejecutar en nueva ventana
start "Streamlit" cmd /k "%LAUNCHER%"

echo -------------------------------------------------------
echo Streamlit se esta iniciando en una nueva ventana.
echo URL: http://localhost:%PORT%
echo Log: %LOGFILE%
echo -------------------------------------------------------
pause

del "%LAUNCHER%" >nul 2>&1
endlocal
popd
exit /b 0
