@echo off
title Sakarya Weekly Reports - Pubblica su GitHub
cd /d "%~dp0"
echo.
echo  Commit e push dei dashboard su GitHub (Pages)...
echo.
git add -A
git commit -m "Update weekly report dashboards"
git push origin main
echo.
echo  Fatto. Pagina: https://fanatics-hue.github.io/sakarya-weekly-reports/
echo.
pause
