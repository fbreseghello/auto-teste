@echo off
cd /d "%~dp0"
if exist "%~dp0AutoTeste.exe" (
  "%~dp0AutoTeste.exe"
  exit /b 0
)
py -m app.main update-app >nul 2>&1
py -m app.gui
