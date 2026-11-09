@echo off
cd /d "%~dp0"
if not exist ".runtime\python.exe" (
  echo First run setup.cmd to install local Python and dependencies.
  pause
  exit /b 1
)
echo DataForge: http://127.0.0.1:8765
echo Open this address in your browser. Keep this window open. Ctrl+C stops the app.
".runtime\python.exe" -m uvicorn dataforge.app:app --host 127.0.0.1 --port 8765
pause
