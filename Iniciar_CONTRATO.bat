@echo off
cd /d "%~dp0"

echo Buscando actualizaciones...
git pull --ff-only
if errorlevel 1 (
    echo No se pudo actualizar automaticamente ^(sin internet, o hay cambios locales sin subir^). Se abre con la version actual.
)

py -m streamlit run Inicio.py
pause
