@echo off
setlocal

cd /d "%~dp0"

set "CODEX_NODE=C:\Users\Defasya\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin"
set "CODEX_BIN=C:\Users\Defasya\.cache\codex-runtimes\codex-primary-runtime\dependencies\bin\fallback"
set "PATH=%CODEX_NODE%;%CODEX_BIN%;%PATH%"

if not exist ".\node_modules\.bin\next.CMD" (
  echo Dependency project belum ditemukan.
  echo Pastikan folder node_modules masih ada / terhubung.
  pause
  exit /b 1
)

echo Starting HSI Dashboard...
echo.
echo Buka di browser:
echo http://localhost:3000
echo.

call ".\node_modules\.bin\next.CMD" dev

pause
