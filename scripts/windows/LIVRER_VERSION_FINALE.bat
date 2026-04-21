@echo off
title DATACOLISA - Version finale
echo.
echo ======================================
echo   CREATION VERSION FINALE DATACOLISA
echo ======================================
echo.
echo Cette commande construit l'exe portable final.
echo Le dossier a livrer sera cree dans code\dist\DATACOLISA_A_NE_PAS_TOUCHER
echo.

cd /d "%~dp0"

powershell -ExecutionPolicy Bypass -File "build_portable.ps1" -ExeName "DATACOLISA"

echo.
echo Version finale terminee.
echo Copie ensuite le dossier code\dist\DATACOLISA_A_NE_PAS_TOUCHER sur l'autre PC.
echo.
pause
