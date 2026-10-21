# ==============================================================================
# Script: Validar-Codificacion-VBA.ps1
# Propósito: Guardián automatizado de codificación e integridad VBA para ResidenciaApp.
#            Invoca el motor de validación Validar-Codificacion-VBA.py.
# ==============================================================================

$scriptPy = Join-Path $PSScriptRoot "Validar-Codificacion-VBA.py"
if (-not (Test-Path $scriptPy)) {
    Write-Error "No se localizó el validador Python en: $scriptPy"
    exit 1
}

& python $scriptPy $args
exit $LASTEXITCODE
