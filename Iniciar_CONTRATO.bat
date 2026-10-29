@echo off
cd /d "%~dp0"

where git >nul 2>nul
if errorlevel 1 (
    echo No se encontro Git ni Python en esta PC todavia.
    echo Antes de usar la app por primera vez, corre Instalar_CONTRATOBOT.bat
    echo ^(esta en la misma carpeta^), que revisa e instala lo que falte.
    pause
    exit /b 1
)

echo Buscando actualizaciones...
git pull --ff-only
if errorlevel 1 (
    echo No se pudo actualizar automaticamente ^(sin internet, o hay cambios locales sin subir^). Se abre con la version actual.
)

rem Activa el chequeo de secretos (gitleaks) antes de cada commit en esta
rem PC -- idempotente, no hace nada si ya estaba activado.
git config core.hooksPath .githooks

py -m streamlit run Inicio.py
pause
