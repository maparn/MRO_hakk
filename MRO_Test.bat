@echo off
chcp 65001 >nul
title Magnetic MRO test

echo Magnetic MRO tootlemine algab.
echo Sulge enne tootlemist koik output Exceli failid.
echo.

wsl.exe -d Ubuntu --cd /home/egors/MRO_hakk -- .venv/bin/python -B scripts/run_pipeline.py --config config.json --fresh --run-tests

if errorlevel 1 (
    echo.
    echo Tootlemine ebaonnestus. Vaata veateadet yleval.
    pause
    exit /b 1
)

echo.
echo Tootlemine valmis. Avan output-kausta.
start "" explorer.exe "\\wsl.localhost\Ubuntu\home\egors\MRO_hakk\excel_files\output"
pause
