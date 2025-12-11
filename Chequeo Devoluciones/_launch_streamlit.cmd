@echo off
echo Iniciando Streamlit en puerto 8501...
echo Comando: streamlit run "C:\Users\Usuario\Downloads\Chequeo Devoluciones\lista.py" --server.port 8501

"streamlit" run "C:\Users\Usuario\Downloads\Chequeo Devoluciones\lista.py" --server.port 8501 >> "C:\Users\Usuario\Downloads\Chequeo Devoluciones\streamlit_run.log" 2>&1

echo.
echo Logs guardados en: "C:\Users\Usuario\Downloads\Chequeo Devoluciones\streamlit_run.log"
echo Presiona una tecla para cerrar esta ventana cuando quieras detener.
pause
