@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  echo Ilk qurulum edilir...
  python -m venv .venv || (echo Python tapilmadi. https://www.python.org/downloads/ & pause & exit /b 1)
  .venv\Scripts\pip install -r requirements.txt
  .venv\Scripts\python -m playwright install chromium
)
set FILE=%~1
if "%FILE%"=="" set /p FILE=Excel faylini bura surukleyin ve Enter basin: 
set FILE=%FILE:"=%
.venv\Scripts\python irshad_print.py "%FILE%" %2 %3 %4 %5 %6
pause
