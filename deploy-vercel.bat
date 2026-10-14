@echo off
setlocal

cd /d "%~dp0"

set "CODEX_NODE=C:\Users\Defasya\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin"
set "CODEX_BIN=C:\Users\Defasya\.cache\codex-runtimes\codex-primary-runtime\dependencies\bin\fallback"
set "PATH=%CODEX_NODE%;%CODEX_BIN%;%PATH%"

echo Checking local production build...
call ".\node_modules\.bin\next.CMD" build
if errorlevel 1 (
  echo.
  echo Build gagal. Perbaiki error di atas dulu sebelum deploy.
  pause
  exit /b 1
)

echo.
echo Login ke Vercel jika diminta.
echo Pastikan environment variables ini sudah diset di Vercel:
echo - DASHBOARD_USERNAME
echo - DASHBOARD_PASSWORD
echo - DASHBOARD_SESSION_SECRET
echo.

call pnpm dlx vercel --prod

pause
