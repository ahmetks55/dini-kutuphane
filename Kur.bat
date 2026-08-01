@echo off
title Dini Kutuphane - Kurulum
cd /d "%~dp0"
echo Dini Kutuphane kuruluyor...
echo.
call npm install
if errorlevel 1 goto fail
node seed.js
if errorlevel 1 goto fail
echo.
echo Kurulum tamamlandi. Simdi "Baslat.bat" dosyasini cift tiklayarak baslatabilirsiniz.
pause
goto end
:fail
echo.
echo Bir hata olustu. Yukaridaki mesaji kontrol edin.
pause
:end
