# Script para crear y activar un entorno virtual, luego instalar los requirements
# Ejecutar desde la raíz del proyecto: .\crear_venv.ps1

$ErrorActionPreference = 'Stop'

Write-Host "Iniciando configuración del entorno virtual..." -ForegroundColor Cyan

# Ruta del directorio del proyecto
$projectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $projectRoot

# Nombre del entorno virtual
$venvDir = ".venv"

if (-Not (Test-Path $venvDir)) {
    Write-Host "Creando entorno virtual en '$venvDir'..." -ForegroundColor Green
    python -m venv $venvDir
} else {
    Write-Host "El entorno virtual ya existe en '$venvDir'." -ForegroundColor Yellow
}

# Activación del entorno virtual
$activateScript = Join-Path $venvDir "Scripts\Activate.ps1"
if (Test-Path $activateScript) {
    Write-Host "Activando el entorno virtual..." -ForegroundColor Green
    & $activateScript
} else {
    Write-Host "No se encontró el script de activación: $activateScript" -ForegroundColor Red
    throw "Fallo al activar el entorno virtual."
}

# Actualizar pip e instalar requirements
Write-Host "Actualizando pip e instalando dependencias..." -ForegroundColor Green
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

Write-Host "Entorno virtual creado e instalado correctamente." -ForegroundColor Cyan
Write-Host "Para usarlo en esta sesión, asegúrate de que el script de activación ya esté ejecutado." -ForegroundColor Cyan
