@echo off
cd /d "%~dp0"
title GrestelPy

:: Recarregar PATH para encontrar uv
set PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%

where uv >nul 2>&1
if errorlevel 1 (
    echo ERRO: uv nao encontrado.
    echo Execute setup_uv.bat primeiro.
    echo.
    pause
    exit /b 1
)

echo Servidor disponivel em: http://localhost:8000
echo Feche esta janela para parar o servidor.
echo.

uv run server.py
pause
