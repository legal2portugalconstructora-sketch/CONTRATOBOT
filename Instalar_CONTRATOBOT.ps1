# Instalar_CONTRATOBOT.ps1 -- Se corre UNA sola vez en una PC nueva, ANTES
# de usar Iniciar_CONTRATO.bat por primera vez. Revisa que esten los
# programas que la app necesita para funcionar (Python, Git, Microsoft
# Excel/Word) y, si falta alguno, muestra el comando (o el link) para
# instalarlo -- no instala nada sin que el usuario lo confirme.
#
# Sin tildes/ñ a proposito en todos los mensajes: PowerShell 5.1 (el que
# trae Windows por defecto) no siempre interpreta bien los acentos al leer
# un script .ps1 guardado en UTF-8 sin BOM, y el texto sale corrupto en la
# consola -- mismo criterio que ya usan los demas .bat de este proyecto.
#
# No se puede correr con doble clic directo (Windows bloquea los .ps1 por
# seguridad) -- por eso existe Instalar_CONTRATOBOT.bat al lado, que SI se
# puede abrir con doble clic y llama a este script.

$ErrorActionPreference = "Continue"

function Escribir-Titulo($texto) {
    Write-Host ""
    Write-Host "=== $texto ===" -ForegroundColor Cyan
}

function Preguntar-SiNo($mensaje) {
    $respuesta = Read-Host "$mensaje (S/N)"
    return $respuesta -match "^[sS]"
}

Escribir-Titulo "ContratoBot -- revisando requisitos de esta PC"

$faltantes = @()

# --- 1) Python (necesario para correr la app y sus dependencias) ---
Write-Host "Verificando Python... " -NoNewline
$pythonOk = $false
try {
    $verPython = & py --version 2>&1
    if ($LASTEXITCODE -eq 0) { $pythonOk = $true }
} catch {}
if (-not $pythonOk) {
    try {
        $verPython = & python --version 2>&1
        if ($LASTEXITCODE -eq 0) { $pythonOk = $true }
    } catch {}
}
if ($pythonOk) {
    Write-Host "OK ($verPython)" -ForegroundColor Green
} else {
    Write-Host "FALTA" -ForegroundColor Red
    $faltantes += [PSCustomObject]@{
        Nombre  = "Python"
        Comando = "winget install --id Python.Python.3.13 -e --source winget --accept-package-agreements --accept-source-agreements"
        Link    = "https://www.python.org/downloads/"
    }
}

# --- 2) Git (necesario para bajar y actualizar la app) ---
Write-Host "Verificando Git... " -NoNewline
$gitOk = $false
try {
    & git --version *> $null
    if ($LASTEXITCODE -eq 0) { $gitOk = $true }
} catch {}
if ($gitOk) {
    Write-Host "OK" -ForegroundColor Green
} else {
    Write-Host "FALTA" -ForegroundColor Red
    $faltantes += [PSCustomObject]@{
        Nombre  = "Git"
        Comando = "winget install --id Git.Git -e --source winget --accept-package-agreements --accept-source-agreements"
        Link    = "https://git-scm.com/download/win"
    }
}

# --- 3) Microsoft Excel / Word reales (la app los necesita instalados y
#         con licencia -- esto NO se puede instalar solo, hace falta
#         comprarlo/activarlo a mano). Se detecta si estan registrados
#         como servidor COM, que es justo lo que la app necesita usar. ---
Write-Host "Verificando Microsoft Excel... " -NoNewline
$excelOk = Test-Path "Registry::HKEY_CLASSES_ROOT\Excel.Application"
if ($excelOk) { Write-Host "OK" -ForegroundColor Green } else { Write-Host "FALTA" -ForegroundColor Red }

Write-Host "Verificando Microsoft Word... " -NoNewline
$wordOk = Test-Path "Registry::HKEY_CLASSES_ROOT\Word.Application"
if ($wordOk) { Write-Host "OK" -ForegroundColor Green } else { Write-Host "FALTA" -ForegroundColor Red }

if (-not $excelOk -or -not $wordOk) {
    $faltantes += [PSCustomObject]@{
        Nombre  = "Microsoft Excel y Word"
        Comando = $null  # no se puede instalar con un comando -- es software con licencia
        Link    = "https://www.microsoft.com/microsoft-365"
    }
}

# --- Si falta algo, avisar claro y ofrecer instalar lo que se pueda ---
if ($faltantes.Count -gt 0) {
    Escribir-Titulo "Faltan instalar estos programas"
    foreach ($item in $faltantes) {
        Write-Host ""
        Write-Host "- $($item.Nombre)" -ForegroundColor Yellow
        if ($item.Comando) {
            Write-Host "  Comando para instalarlo:"
            Write-Host "  $($item.Comando)" -ForegroundColor White
            if (Preguntar-SiNo "  Instalarlo ahora?") {
                Invoke-Expression $item.Comando
            }
        } else {
            Write-Host "  Este no se instala con un comando -- es software con licencia propia."
            Write-Host "  Consiguelo/activalo en: $($item.Link)"
        }
    }
    Write-Host ""
    Write-Host "Volve a correr este instalador despues de instalar lo que falte," -ForegroundColor Cyan
    Write-Host "para confirmar que ya quedo todo listo." -ForegroundColor Cyan
    Read-Host "`nPresiona Enter para salir"
    exit 1
}

Write-Host ""
Write-Host "Todos los programas necesarios ya estan instalados." -ForegroundColor Green

# --- 4) Es esta carpeta un repositorio de git? (hace falta para que
#         Iniciar_CONTRATO.bat pueda actualizarse solo con git pull) ---
Escribir-Titulo "Verificando la carpeta del proyecto"
$esRepoGit = Test-Path ".git"
if (-not $esRepoGit) {
    Write-Host "Esta carpeta NO es un repositorio de git -- probablemente se bajo como" -ForegroundColor Yellow
    Write-Host "un .zip en vez de clonarse. Asi, Iniciar_CONTRATO.bat no va a poder" -ForegroundColor Yellow
    Write-Host "actualizarse solo. Recomendado: borrar esta carpeta y clonarla de nuevo con:" -ForegroundColor Yellow
    Write-Host "  git clone <URL del repositorio>" -ForegroundColor White
} else {
    Write-Host "OK -- esta carpeta es un repositorio de git, las actualizaciones automaticas van a funcionar." -ForegroundColor Green
}

# --- 5) Instalar las librerias de Python que la app necesita ---
Escribir-Titulo "Instalando las librerias de Python (requirements.txt)"
py -m pip install -r requirements.txt

# --- 6) Activar el chequeo de secretos (gitleaks) antes de cada commit --
#         en esta PC no hace falta si nadie va a tocar codigo aca, pero no
#         esta de mas dejarlo activado (no molesta si gitleaks no esta
#         instalado, ver .githooks/pre-commit). ---
if ($esRepoGit) {
    git config core.hooksPath .githooks
}

# --- 7) Recordatorio de lo que NO viaja con git (datos del negocio) ---
Escribir-Titulo "Ultimos pasos -- esto NO se baja solo de GitHub"
Write-Host "Estas carpetas/archivos tienen datos del negocio y hay que copiarlos a mano"
Write-Host "(pediselos a Jesus/Portugal y Rivera -- por ejemplo por USB o un Drive compartido):"
Write-Host "  - La carpeta PLANTILLAS\ completa (las plantillas de cada proyecto)"
Write-Host "  - El archivo config.yaml (si la app vuelve a pedir usuario y contrasena)"
Write-Host ""
Write-Host "Una vez que esten esas 2 cosas en su lugar, ya podes abrir la app con" -ForegroundColor Green
Write-Host "Iniciar_CONTRATO.bat." -ForegroundColor Green

Read-Host "`nPresiona Enter para salir"
