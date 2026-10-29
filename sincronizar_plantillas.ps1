# sincronizar_plantillas.ps1 -- Sube a GitHub los cambios del dia en
# PLANTILLAS/ (y solo esa carpeta -- el codigo se sigue subiendo a mano,
# via Claude, como siempre). Pensado para correr solo, todos los dias a
# las 18:00, desde el Programador de tareas de Windows (ver
# Instalar_Sincronizacion_Plantillas.ps1 para registrar esa tarea).
#
# Si no hay ningun cambio en PLANTILLAS/ ese dia, no hace ningun commit
# (evita commits vacios). Sin tildes/ñ en los mensajes, mismo motivo que
# el resto de los scripts de este proyecto (PowerShell 5.1 + consola).

$ErrorActionPreference = "Continue"
Set-Location -Path $PSScriptRoot

$fecha = Get-Date -Format "yyyy-MM-dd HH:mm"
$log = Join-Path $PSScriptRoot "sincronizar_plantillas.log"

function Escribir-Log($texto) {
    $linea = "[$fecha] $texto"
    Write-Host $linea
    Add-Content -Path $log -Value $linea
}

Escribir-Log "--- Sincronizando PLANTILLAS/ ---"

# Traer primero lo que haya cambiado desde otra PC, para no pisar nada.
git pull --ff-only *> $null
if ($LASTEXITCODE -ne 0) {
    Escribir-Log "AVISO: no se pudo actualizar antes de sincronizar (sin internet, o hay cambios locales sin subir). Se sigue igual."
}

git add PLANTILLAS/

# Si no quedo nada en el stage, no hay nada nuevo que subir -- cortar aca.
git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Escribir-Log "Sin cambios en PLANTILLAS/ hoy -- no se hizo ningun commit."
    exit 0
}

git commit -m "Actualizacion automatica de PLANTILLAS/ ($fecha)" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Escribir-Log "ERROR: el commit automatico fallo (revisa sincronizar_plantillas.log y corre 'git status' a mano)."
    exit 1
}

git push origin master | Out-Null
if ($LASTEXITCODE -ne 0) {
    Escribir-Log "ERROR: el commit se creo pero no se pudo subir a GitHub (revisa la conexion a internet). Se va a reintentar al otro dia, o corre 'git push' a mano."
    exit 1
}

Escribir-Log "Listo -- cambios de PLANTILLAS/ subidos a GitHub."
