@echo off
rem ContratoBot_Instalador.bat -- UN SOLO archivo para instalar la app en
rem una PC nueva de punta a punta: baja el proyecto de GitHub, revisa e
rem instala lo que haga falta (Python, Git si no lo tenes, etc.) y deja un
rem acceso directo en el Escritorio. Se corre con doble clic, sin escribir
rem ningun comando.
rem
rem Este archivo NO necesita estar adentro de la carpeta del proyecto --
rem se baja suelto (ej. desde la pagina de Releases de GitHub) y se corre
rem desde donde sea (Descargas, Escritorio, etc.).

setlocal
set "REPO_URL=https://github.com/jediri2010-web/CONTRATOBOT.git"
set "DESTINO=%USERPROFILE%\Desktop\CONTRATOBOT"

echo ===============================================
echo   Instalador de ContratoBot
echo ===============================================
echo.

where git >nul 2>nul
if errorlevel 1 (
    echo No se encontro Git en esta PC -- hace falta para bajar y
    echo actualizar la app. Instalalo desde:
    echo   https://git-scm.com/download/win
    echo.
    echo Despues de instalarlo, volve a correr este instalador.
    pause
    exit /b 1
)

if exist "%DESTINO%\.git" (
    echo Ya existe una instalacion en "%DESTINO%" -- actualizando en vez
    echo de bajar todo de nuevo.
    pushd "%DESTINO%"
    git pull --ff-only
    popd
) else (
    echo Bajando ContratoBot a "%DESTINO%"...
    echo ^(Si es la primera vez, se te va a abrir el navegador para que
    echo inicies sesion en GitHub -- iniciala con tu propia cuenta.^)
    echo.
    git clone "%REPO_URL%" "%DESTINO%"
    if errorlevel 1 (
        echo.
        echo No se pudo bajar el proyecto. Revisa que tengas acceso al
        echo repositorio ^(solo lectura^) y volve a intentar.
        pause
        exit /b 1
    )
)

echo.
echo ===============================================
echo   Revisando requisitos ^(Python, Excel, Word^)...
echo ===============================================
powershell -NoProfile -ExecutionPolicy Bypass -File "%DESTINO%\Instalar_CONTRATOBOT.ps1"

echo.
echo Creando acceso directo en el Escritorio...
powershell -NoProfile -ExecutionPolicy Bypass -Command ^
    "$ws = New-Object -ComObject WScript.Shell; " ^
    "$lnk = $ws.CreateShortcut('%USERPROFILE%\Desktop\ContratoBot.lnk'); " ^
    "$lnk.TargetPath = '%DESTINO%\Iniciar_CONTRATO.bat'; " ^
    "$lnk.WorkingDirectory = '%DESTINO%'; " ^
    "$lnk.IconLocation = '%DESTINO%\CONTRATOBOT.ico'; " ^
    "$lnk.Save()"

echo.
echo ===============================================
echo   Listo. Ya tenes un acceso directo "ContratoBot"
echo   en el Escritorio para abrir la app.
echo ===============================================
pause
