@echo off
setlocal
cd /d "%~dp0"

if not exist "node_modules" (
  echo Installing dependencies...
  call npm install
  if errorlevel 1 (
    echo.
    echo npm install failed.
    pause
    exit /b 1
  )
)

echo Starting LED Cabling Web App...
echo Your browser should open automatically.
call npm run dev -- --host --open

if errorlevel 1 (
  echo.
  echo The local server stopped with an error.
  pause
)
