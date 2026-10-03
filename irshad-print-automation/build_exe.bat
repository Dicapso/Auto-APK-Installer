@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Ara uzu tek exe faylina cevirir (kod gizlenir).
rem Filial kilidi:  build_exe.bat irshadecemi2@gmail.com   -> dist\IrshadCap_irshadecemi2.exe
rem Kilidsiz:       build_exe.bat                           -> dist\IrshadCap.exe
if not exist .venv (
  python -m venv .venv || (echo Python tapilmadi. & pause & exit /b 1)
)
.venv\Scripts\pip install -q -r requirements.txt pyinstaller
set EMAIL=%~1
if "%EMAIL%"=="" set /p EMAIL=Filial e-poctu (bos = kilidsiz): 
set NAME=IrshadCap
if not "%EMAIL%"=="" (
  for /f "tokens=1 delims=@" %%a in ("%EMAIL%") do set NAME=IrshadCap_%%a
  > branch_lock.py echo ALLOWED_EMAILS = ["%EMAIL%"]
) else (
  if exist branch_lock.py del branch_lock.py
)
.venv\Scripts\pyinstaller --noconfirm --onefile --windowed --name %NAME% --hidden-import branch_lock --collect-all playwright --collect-all customtkinter irshad_gui.py
if exist branch_lock.py del branch_lock.py
echo.
echo Hazirdir: %~dp0dist\%NAME%.exe
pause
