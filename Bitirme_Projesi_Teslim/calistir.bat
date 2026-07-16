@echo off
setlocal EnableExtensions
chcp 65001 >nul
cd /d "%~dp0"
title TBMYO Analiz - Kurulum ve Calistirma

echo ==========================================================
echo   TBMYO IKI DONEM KARSILASTIRMALI ANALIZ
echo ==========================================================
echo.

rem Once Windows Python Launcher, sonra python komutu kontrol edilir.
set "PY_CMD="
py -3 --version >nul 2>&1
if not errorlevel 1 set "PY_CMD=py -3"

if not defined PY_CMD (
    python --version >nul 2>&1
    if not errorlevel 1 set "PY_CMD=python"
)

if not defined PY_CMD goto python_yok

echo [1/4] Python bulundu.
%PY_CMD% --version

rem Kutuphaneleri bilgisayarin geneline degil, proje icindeki sanal ortama kur.
if not exist ".venv\Scripts\python.exe" (
    echo.
    echo [2/4] Ilk kurulum yapiliyor: .venv olusturuluyor...
    %PY_CMD% -m venv .venv
    if errorlevel 1 goto hata
) else (
    echo [2/4] Mevcut .venv kullaniliyor.
)

set "VENV_PY=%~dp0.venv\Scripts\python.exe"

echo.
echo [3/4] Gerekli kutuphaneler kontrol ediliyor...
"%VENV_PY%" -m pip install --disable-pip-version-check -r requirements.txt
if errorlevel 1 goto pip_hata

echo.
echo [4/4] Analiz baslatiliyor...
"%VENV_PY%" bitirme_projesi_analiz.py
if errorlevel 1 goto hata

echo.
echo ==========================================================
echo   ISLEM BASARIYLA TAMAMLANDI
echo ==========================================================
echo Sonuclar klasoru aciliyor...
start "" "%~dp0Sonuclar"
pause
exit /b 0

:python_yok
echo.
echo ==========================================================
echo   PYTHON BULUNAMADI
echo ==========================================================
echo.
echo 1. Acilan resmi Python sayfasindan Windows 64-bit Python kurun.
echo 2. Kurulumda "Add Python to PATH" secenegini isaretleyin.
echo 3. Kurulum bitince bu pencereyi kapatin.
echo 4. calistir.bat dosyasina yeniden cift tiklayin.
echo.
echo Not: Python 3.13 64-bit ile proje test edilmistir.
start "" "https://www.python.org/downloads/windows/"
pause
exit /b 1

:pip_hata
echo.
echo ==========================================================
echo   KUTUPHANE KURULUMU BASARISIZ
echo ==========================================================
echo Internet baglantisini kontrol edin ve tekrar calistirin.
echo Kurumsal/ag engeli varsa CMD'yi normal kullanici olarak yeniden deneyin.
pause
exit /b 1

:hata
echo.
echo ==========================================================
echo   ISLEM BASARISIZ
echo ==========================================================
echo Yukaridaki son hata mesajinin ekran goruntusunu alin.
pause
exit /b 1
