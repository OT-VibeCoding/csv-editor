@echo off
rem CsvEditor build script (uses the C# compiler included in Windows / .NET Framework 4.x)
cd /d "%~dp0"
set "CSC=%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\csc.exe"
if not exist "%CSC%" set "CSC=%WINDIR%\Microsoft.NET\Framework\v4.0.30319\csc.exe"
if not exist "%CSC%" (
  echo csc.exe was not found. Please enable .NET Framework 4.8.
  pause
  exit /b 1
)
"%CSC%" /nologo /target:winexe /optimize+ /codepage:65001 /out:CsvEditor.exe /r:System.Windows.Forms.dll /r:System.Drawing.dll /r:System.Core.dll CsvEditor.cs
if errorlevel 1 (
  echo Build FAILED.
  pause
  exit /b 1
)
echo.
echo Build OK: CsvEditor.exe
pause
