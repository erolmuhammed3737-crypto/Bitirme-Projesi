@echo off
cd /d "%~dp0"
if exist "Sonuclar" rmdir /s /q "Sonuclar"
mkdir "Sonuclar"
echo Sonuclar klasoru temizlendi.
pause
