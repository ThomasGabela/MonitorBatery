@echo off
rem Compila MonitorBateria.exe con el compilador de C# que ya trae Windows (.NET Framework 4).
cd /d "%~dp0"
set CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if not exist "%CSC%" (
  echo No se encontro csc.exe en %CSC%
  pause
  exit /b 1
)
"%CSC%" /nologo /optimize+ /platform:x64 /target:winexe /r:System.Core.dll /r:System.Management.dll /r:System.Windows.Forms.dll /r:System.Drawing.dll /out:MonitorBateria.exe MonitorBateria.cs
if errorlevel 1 (
  echo Error al compilar.
  pause
  exit /b 1
)
echo Listo: MonitorBateria.exe
pause
