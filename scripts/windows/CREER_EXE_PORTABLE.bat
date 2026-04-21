@echo off
title Build DATACOLISA Portable
echo.
echo ======================================
echo   BUILD DATACOLISA - EXE PORTABLE
echo ======================================
echo.

cd /d "%~dp0"

powershell -ExecutionPolicy Bypass -File "build_portable.ps1"

echo.
echo Appuie sur une touche pour fermer...
pause > nul
