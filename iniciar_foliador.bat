@echo off
title Foliador Automatico PAC / POS - ColDev
cd /d "%~dp0"

echo ======================================================
echo    COLDEV - FOLIADOR AUTOMATICO DE PLANILLAS PAC/POS
echo ======================================================
echo Iniciando interfaz grafica...

python app_gui.py

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Hubo un problema al iniciar la aplicacion.
    echo Verifique que Python este instalado en su sistema.
    pause
)
