@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================================
echo  PIPELINE COMPLETO DE MATURIDADE BIM
echo ============================================================
echo.

echo [1/3] Extraindo metricas dos modelos RVT
echo       (~30 min para 11 modelos - Revit abre/fecha para cada um)
echo.
python run_batch.py
if errorlevel 1 (
    echo.
    echo [ERRO] run_batch.py falhou. Verifique batch.log
    pause
    exit /b 1
)

echo.
echo [2/3] Consolidando em Excel
python consolidar.py
if errorlevel 1 (
    echo.
    echo [ERRO] consolidar.py falhou.
    pause
    exit /b 1
)

echo.
echo [3/3] Abrindo dashboard
echo       (Ctrl+C para encerrar o Streamlit)
echo.
python -m streamlit run dashboard.py

pause