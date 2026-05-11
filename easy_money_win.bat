@echo off
setlocal
set "ROOT=%~dp0"
set "PY=%ROOT%.venv\Scripts\python.exe"
if not exist "%PY%" (
  echo Error: virtual environment not found: %PY%
  echo Run: "%ROOT%.python\python.exe" -m venv "%ROOT%.venv"
  exit /b 1
)
"%PY%" "%ROOT%easy_money_win.py" %*
