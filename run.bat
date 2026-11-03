@echo off
title Sistem Monitor (SiPaMeL)
cd /d "%~dp0"

echo ============================================================
echo   SISTEM MONITORING STASIUN (SiPaMeL)
echo ============================================================
echo.

set PYTHONUTF8=1

if not exist "venv\Scripts\python.exe" (
    echo [ERROR] Virtual environment 'venv' tidak ditemukan!
    echo Pastikan folder venv tersedia di direktori ini.
    pause
    exit /b 1
)

echo [1/2] Membuka tampilan aplikasi di browser...
start "" "http://localhost:5000"

echo [2/2] Menjalankan server aplikasi...
echo Server aktif pada: http://localhost:5000
echo Tekan CTRL + C di jendela ini untuk menghentikan server.
echo ============================================================
echo.

"venv\Scripts\python.exe" app.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Aplikasi berhenti dengan kode error: %ERRORLEVEL%
    pause
)
