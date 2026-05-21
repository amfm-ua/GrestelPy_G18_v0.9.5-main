@echo off
cd /d "%~dp0"
title GrestelPy - Configuracao (uv)
echo ============================================
echo  GrestelPy - Configuracao Inicial (uv)
echo ============================================
echo.

:: Verificar se uv ja esta instalado
where uv >nul 2>&1
if not errorlevel 1 (
    echo uv ja encontrado.
    goto sync
)

echo A instalar uv...
powershell -Command "irm https://astral.sh/uv/install.ps1 | iex"
if errorlevel 1 (
    echo.
    echo ERRO: Falha ao instalar uv.
    echo Verifique a ligacao a internet e tente novamente.
    pause
    exit /b 1
)

:: Recarregar PATH para encontrar uv
set PATH=%USERPROFILE%\.local\bin;%USERPROFILE%\.cargo\bin;%PATH%

:sync
echo A instalar dependencias...
uv sync
if errorlevel 1 (
    echo.
    echo ERRO: Falha ao instalar dependencias.
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Configuracao concluida!
echo  Execute start_uv.bat para abrir o GrestelPy.
echo ============================================
echo.
pause
