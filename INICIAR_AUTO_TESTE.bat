@echo off
cd /d "%~dp0"
py -m app.main update-app >nul 2>&1
py -m app.gui
