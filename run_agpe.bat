@echo off
cd /d "%~dp0"

REM Usar Python del entorno virtual si existe
if exist venv\Scripts\python.exe (
    echo Ejecutando con Python del entorno virtual...
    venv\Scripts\python.exe main.py
) else (
    echo Ejecutando con Python del sistema...
    python main.py
)

pause
