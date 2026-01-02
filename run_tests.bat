@echo off
cls
echo ==========================================
echo   METEO CANTIERE - TEST SUITE RUNNER
echo ==========================================

if not exist ".venv" (
    echo [ERRORE] Ambiente virtuale non trovato. Avvia prima Avvia_Meteo.bat
    pause
    exit /b
)

call .venv\Scripts\activate.bat

echo [INFO] Installazione tool di testing...
pip install pytest pytest-cov pytest-mock --quiet

echo [INFO] Esecuzione Test & Calcolo Copertura...
echo.

:: Esegue i test e calcola la copertura ignorando le librerie di sistema e i file di test stessi
pytest tests/ -v --cov=Meteo_Lavoro --cov-report=term-missing

echo.
echo ==========================================
echo   TEST COMPLETATI
echo ==========================================
pause
