@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  echo Ilk qurulum edilir...
  python -m venv .venv || (echo Python tapilmadi. https://www.python.org/downloads/ & pause & exit /b 1)
  .venv\Scripts\pip install -r requirements.txt
)
echo.
echo  1 - Excel-deki mehsullari cap ucun sec (esas)
echo  2 - Excel-deki mehsullari siyahidan sil
echo.
set /p MODE=Secim (1/2, bos = 1): 
set FILE=%~1
if "%FILE%"=="" set /p FILE=Excel faylini bura surukleyin ve Enter basin: 
set FILE=%FILE:"=%
set EXTRA=
if "%MODE%"=="2" set EXTRA=--remove
.venv\Scripts\python irshad_print.py "%FILE%" %EXTRA% %2 %3 %4 %5 %6
pause
