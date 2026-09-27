@echo off
chcp 65001 >nul
cd /d "%~dp0"
rem Excel-dəki ID-ləri çap siyahısından silir (qırmızı − düyməsi)
set FILE=%~1
if "%FILE%"=="" set /p FILE=Excel faylini bura surukleyin ve Enter basin: 
set FILE=%FILE:"=%
call run.bat "%FILE%" --remove %2 %3 %4
