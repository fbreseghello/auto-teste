@echo off
cd /d "%~dp0"
py -m app.main --db-path data\aurha_teste.db sync-yampi --client aurha_artesanatoholistico --start-date 31/01/2026 --end-date 13/02/2026
py -m app.main --db-path data\aurha_teste.db export-monthly --client aurha_artesanatoholistico --start-date 31/01/2026 --end-date 13/02/2026 --output exports\aurha_artesanatoholistico_mensal.csv
echo.
echo Teste finalizado. Verifique o arquivo exports\aurha_artesanatoholistico_mensal.csv
pause
