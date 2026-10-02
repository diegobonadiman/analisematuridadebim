@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo Reconsolidando Excel a partir dos JSONs existentes...
echo (util quando voce muda a regua de notas)
echo.
python consolidar.py
echo.
pause