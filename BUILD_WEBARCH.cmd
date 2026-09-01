@echo off
setlocal
cd /d "%~dp0"
echo ============================================================
echo  WEBARCH v1.1 - BUILD DEL WORKBOOK OPERATIVO
echo ============================================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0BUILD\Build-WebArch.ps1"
set EXITCODE=%ERRORLEVEL%
echo.
if not "%EXITCODE%"=="0" (
  echo Build non completata. Consultare BUILD\Build-WebArch.log
) else (
  echo Build completata.
)
echo.
pause
exit /b %EXITCODE%
