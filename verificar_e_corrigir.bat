@echo off
chcp 65001 >nul
title Klassmatt - Verificar e Corrigir

echo ======================================
echo  ETAPA 1 - Verificacao (skip ja verificados)
echo ======================================
echo.
python verify_items.py --skip-verified
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] verify_items.py falhou!
    pause
    exit /b 1
)

echo.
echo ======================================
echo  ETAPA 2 - Correcao (divergentes do report)
echo ======================================
echo.
python fix_items.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] fix_items.py falhou!
    pause
    exit /b 1
)

echo.
echo ======================================
echo  ETAPA 3 - Re-verificacao dos corrigidos
echo ======================================
echo.
python verify_items.py --only-divergent
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERRO] verify_items.py falhou!
    pause
    exit /b 1
)

echo.
echo ======================================
echo  CONCLUIDO
echo ======================================
pause
