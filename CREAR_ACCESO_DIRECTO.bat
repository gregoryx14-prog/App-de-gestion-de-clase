@echo off
setlocal
set "TARGET=%~dp0INICIAR_GESTION_CLASE.bat"
set "SHORTCUT=%USERPROFILE%\Desktop\Gestion de Clase V5.1.lnk"
powershell -NoProfile -ExecutionPolicy Bypass -Command "$ws=New-Object -ComObject WScript.Shell; $s=$ws.CreateShortcut('%SHORTCUT%'); $s.TargetPath='%TARGET%'; $s.WorkingDirectory='%~dp0'; $s.IconLocation='%SystemRoot%\System32\SHELL32.dll,220'; $s.Description='Iniciar Gestion de Clase V5.1'; $s.Save()"
if exist "%SHORTCUT%" (echo Acceso directo creado en el Escritorio.) else (echo No se pudo crear el acceso directo.)
timeout /t 2 /nobreak >nul
