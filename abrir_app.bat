@echo off
cd /d "%~dp0"
if not "%AUTO_TESTE_GITHUB_REPO%"=="" (
  py -m app.main update-app
)
py -m app.gui
