@echo off
title Cloudflare Tunnel (SiPaMeL Public Link)
cd /d "%~dp0"

echo ============================================================
echo   MEMBUAT LINK PUBLIK / TUNNEL (CLOUDFLARE)
echo ============================================================
echo.
echo Pastikan aplikasi SiPaMeL (run.bat) SUDAH BERJALAN terlebih dahulu!
echo Menghubungkan ke Cloudflare...
echo Link publik HTTPS (trycloudflare.com) akan tampil di bawah:
echo ============================================================
echo.

cloudflared.exe tunnel --url http://localhost:5000

pause
