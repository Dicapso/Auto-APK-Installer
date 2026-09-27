@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Skripti tek IrshadCap.exe faylina cevirir (kod gizlenir). Netice: dist\IrshadCap.exe
if not exist .venv (
  python -m venv .venv || (echo Python tapilmadi. & pause & exit /b 1)
  .venv\Scripts\pip install -r requirements.txt
)
.venv\Scripts\pip install pyinstaller
.venv\Scripts\pyinstaller --noconfirm --onefile --name IrshadCap --collect-all playwright irshad_print.py
echo.
echo Hazirdir: %~dp0dist\IrshadCap.exe
pause
