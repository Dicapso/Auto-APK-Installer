@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Ara uzu tek IrshadCap.exe faylina cevirir (kod gizlenir). Netice: dist\IrshadCap.exe
if not exist .venv (
  python -m venv .venv || (echo Python tapilmadi. & pause & exit /b 1)
)
.venv\Scripts\pip install -q -r requirements.txt pyinstaller
.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name IrshadCap --collect-all playwright --collect-all customtkinter irshad_gui.py
echo.
echo Hazirdir: %~dp0dist\IrshadCap.exe
pause
