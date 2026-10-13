@echo off
cd /d "%~dp0"

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
