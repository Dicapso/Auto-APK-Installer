@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (
  echo Ilk qurulum edilir...
  python -m venv .venv || (echo Python tapilmadi. https://www.python.org/downloads/ & pause & exit /b 1)
)
rem Yeni paketler (ara uz ucun) elave olunubsa, qurasdir
.venv\Scripts\pip install -q -r requirements.txt
start "" .venv\Scripts\pythonw irshad_gui.py %1
