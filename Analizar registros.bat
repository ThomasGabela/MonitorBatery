@echo off
rem Genera el informe del monitor (por defecto, las sesiones de las ultimas 24 horas) y lo abre.
rem Otros periodos:  "Analizar registros.bat" -Horas 72    o    "Analizar registros.bat" -Todas
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0analizar.ps1" %*
if errorlevel 1 pause
