@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Skript brauzerindeki köhne/boyuk çap siyahisini silir (giriş yeniden soruşulur)
set FILE=%~1
if "%FILE%"=="" set /p FILE=Excel faylini bura surukleyin ve Enter basin: 
set FILE=%FILE:"=%
call run.bat "%FILE%" --reset %2 %3 %4
