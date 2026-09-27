@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv (echo Evvelce run.bat ile qurulum edin. & pause & exit /b 1)
.venv\Scripts\python irshad_print.py - --diagnose
pause
