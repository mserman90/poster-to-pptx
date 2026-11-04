@echo off
title Poster2PPTX - Akilli Afis Donusturucu
cd /d "%~dp0"
echo Poster2PPTX baslatiliyor...
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"
pause
