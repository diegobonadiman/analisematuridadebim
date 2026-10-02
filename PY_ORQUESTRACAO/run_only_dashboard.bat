@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Abrindo dashboard com o Excel atual...
echo (Ctrl+C para encerrar)
echo.
python -m streamlit run dashboard.py
pause