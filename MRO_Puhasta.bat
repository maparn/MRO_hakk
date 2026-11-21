@echo off
chcp 65001 >nul
title Magnetic MRO puhastamine

echo See eemaldab seadistatud input- ja output-Exceli failid.
echo Kasuta seda ainult ENNE uute sisendfailide laadimist.
choice /C JN /M "Kas soovid jatkata"
if errorlevel 2 exit /b 0

wsl.exe -d Ubuntu --cd /home/egors/MRO_hakk -- .venv/bin/python -B scripts/run_pipeline.py --config config.json --clean-all

if errorlevel 1 (
    echo Puhastamine ebaonnestus.
    pause
    exit /b 1
)

echo Puhastamine valmis. Laadi uued failid excel_files\input kausta.
pause
