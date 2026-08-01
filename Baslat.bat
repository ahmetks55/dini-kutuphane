@echo off
title Dini Kutuphane
cd /d "%~dp0"
echo Dini Kutuphane baslatiliyor...
echo.
echo Kapatmak icin bu pencereyi kapatabilirsiniz.
start "" http://localhost:3000
node server.js
pause
