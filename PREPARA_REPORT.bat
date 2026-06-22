@echo off
title Sakarya Weekly Reports - Prepara Report
cd /d "%~dp0"
echo.
echo  Scansione cartelle, estrazione ultimo .doc e allineamento nomi/index...
echo.
python prepara_report.py
echo.
pause
