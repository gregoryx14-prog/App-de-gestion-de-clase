@echo off
setlocal
cd /d "%~dp0"
title Gestion de Clase V5.1 - Detener
echo Deteniendo Gestion de Clase V5.1...
for /f "tokens=5" %%P in ('netstat -ano ^| findstr ":8000" ^| findstr "LISTENING"') do taskkill /PID %%P /F >nul 2>&1
for /f "tokens=2" %%P in ('tasklist ^| findstr /I "uvicorn.exe"') do taskkill /PID %%P /F >nul 2>&1
echo Sistema detenido.
pause
