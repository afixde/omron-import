@echo off
copy "data\excel\Omron - Kopie.xlsx" data\excel\Omron.xlsx
rem pause
python src\main.py
pause