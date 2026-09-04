@echo off
title SGA Label Studio - Visor de Etiquetas Fitosanitarias
echo ========================================================
echo   Iniciando SGA Label Studio (Generador de Etiquetas)
echo ========================================================
echo.
cd /d "%~dp0"
start http://localhost:8000
python -m uvicorn app:app --host 127.0.0.1 --port 8000
pause
