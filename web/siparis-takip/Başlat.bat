@echo off
chcp 65001 >nul
title Sipariş Takip Uygulaması
cd /d "G:\Drive'ım\SİPARİŞ TAKİP"

where python >nul 2>&1
if errorlevel 1 (
  echo Python bulunamadi! Lutfen Python'i kurun (python.org)
  pause
  exit /b 1
)

echo Sipariş Takip sunucusu başlatılıyor...
start "SiparisSunucu" python main.py

REM Sunucu hazır olana kadar bekle (HTTP 200 gelene kadar dene)
set /a n=0
:bekle
timeout /t 1 /nobreak >nul
curl -s -o nul -w "%%{http_code}" http://127.0.0.1:8000/ | findstr "200" >nul && goto ac
set /a n+=1
if %n% lss 15 goto bekle

:ac
start "" "http://127.0.0.1:8000"
echo.
echo ========================================
echo  UYGULAMA AÇILDI: http://127.0.0.1:8000
echo  Veriler siparis.xlsx icine kaydedilir.
echo  Kapatmak icin bu pencereyi KAPATIN.
echo ========================================
echo.

REM Kullanıcı ENTER'a basana kadar bekle
pause >nul

REM Sunucu penceresini kapat
taskkill /fi "WINDOWTITLE eq SiparisSunucu*" >nul 2>&1
taskkill /f /im python.exe >nul 2>&1
