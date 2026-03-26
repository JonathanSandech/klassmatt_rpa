@echo off
setlocal enabledelayedexpansion

set MAX_RESTARTS=100
set RESTART_DELAY=15
set restart_count=0

:loop
if !restart_count! geq %MAX_RESTARTS% (
    echo [MONITOR] Max restarts %MAX_RESTARTS% atingido. Encerrando.
    goto :end
)

set /a attempt=restart_count+1
echo ==========================================
echo [MONITOR] Iniciando verify_and_fix.py - tentativa !attempt!
echo [MONITOR] Args: %*
echo [MONITOR] %date% %time%
echo ==========================================

echo [MONITOR] Matando processos do browser...
taskkill /F /IM chrome.exe /T >nul 2>&1
taskkill /F /IM chromium.exe /T >nul 2>&1
timeout /t 3 /nobreak >nul

python verify_and_fix.py %*
set exit_code=!errorlevel!

echo.
echo [MONITOR] verify_and_fix.py saiu com codigo: !exit_code! em %date% %time%

if !exit_code! equ 0 (
    echo [MONITOR] verify_and_fix.py finalizou com sucesso. Encerrando.
    goto :end
)

set /a restart_count+=1
echo [MONITOR] Matando processos do browser...
taskkill /F /IM chrome.exe /T >nul 2>&1
taskkill /F /IM chromium.exe /T >nul 2>&1
echo [MONITOR] Aguardando %RESTART_DELAY%s antes de reiniciar...
timeout /t %RESTART_DELAY% /nobreak >nul
goto :loop

:end
echo [MONITOR] Fim. Total de restarts: !restart_count!
pause
