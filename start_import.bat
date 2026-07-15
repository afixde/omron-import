@echo off
copy "data\excel\Omron - Kopie.xlsx" data\excel\Omron.xlsx
copy "data\archive\Ihr angeforderter OMRON-Bericht vom 13 Juli 2026 bis zum 13 Juli 2026.csv" data\csv
rem pause
python src\main.py
pause