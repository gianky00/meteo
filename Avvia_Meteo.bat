@echo off
title Meteo Cantiere Launcher
cls
echo ===================================================
echo   SUPER METEO CANTIERE v4.0 (DB + Safety)
echo ===================================================

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERRORE] Python non trovato!
    pause
    exit /b
)

if not exist ".venv" (
    echo [INFO] Configurazione ambiente virtuale...
    python -m venv .venv
)

echo [INFO] Caricamento librerie (incluso Grafici)...
call .venv\Scripts\activate.bat
pip install -r requirements.txt --quiet --disable-pip-version-check

echo [INFO] Avvio applicazione...
start /B pythonw Meteo_Lavoro.py
exit
