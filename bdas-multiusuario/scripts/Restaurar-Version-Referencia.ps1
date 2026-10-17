# ==============================================================================
# Script: Restaurar-Version-Referencia.ps1
# Propósito: Restaurar de forma instantánea la Versión de Referencia Estable
#            (Hito 2026-10-17: Selector Multirresidencia, LOG, RibbonX Dinámico,
#             Tarifas Dinámicas Anti-Solapes y Blindaje Universal de Fórmulas)
#            en caso de que se produzcan errores en futuras modificaciones.
# ==============================================================================

param(
    [switch]$Force
)

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "   RESTAURACION DE VERSION DE REFERENCIA (HITO 2026-10-17)       " -ForegroundColor Cyan
Write-Host "   (Multirresidencia + LOG + RibbonX + Blindaje de Formulas)     " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

$refDir = "H:\ResidenciaBD\Plantilla\VersionReferencia"
$refFileXlsm = Join-Path $refDir "BDAS_Multiusuario_VERSION_REFERENCIA.xlsm"
$refFileAccdb = Join-Path $refDir "Residencia_BE_VERSION_REFERENCIA.accdb"

$activeXlsm = "H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm"
$deployXlsm = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"
$activeAccdb = "H:\ResidenciaBD\Residencia_BE.accdb"

if (-not (Test-Path $refFileXlsm)) {
    Write-Host "ERROR: No se encuentra la copia protegida en $refFileXlsm" -ForegroundColor Red
    exit 1
}

# 1. Crear backup de seguridad del estado actual antes de sobreescribir
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$backupDir = "H:\ResidenciaBD\BackupsDiarios"

if (Test-Path $activeXlsm) {
    $bkpXlsm = Join-Path $backupDir "BDAS_Multiusuario_PREVIO_RESTAURACION_$timestamp.xlsm"
    Copy-Item -LiteralPath $activeXlsm -Destination $bkpXlsm -Force
    Write-Host "[INFO] Copia de seguridad previa guardada en: $bkpXlsm" -ForegroundColor Gray
}

# 2. Restaurar el libro Excel en el repositorio de trabajo
Copy-Item -LiteralPath $refFileXlsm -Destination $activeXlsm -Force
# Quitar atributo de solo lectura en la copia de trabajo
Set-ItemProperty -LiteralPath $activeXlsm -Name IsReadOnly -Value $false
Write-Host "[OK] Libro Excel restaurado con exito: $activeXlsm" -ForegroundColor Green

# 3. Restaurar en H:\ResidenciaBD si existe
if (Test-Path $deployXlsm) {
    Copy-Item -LiteralPath $refFileXlsm -Destination $deployXlsm -Force
    Set-ItemProperty -LiteralPath $deployXlsm -Name IsReadOnly -Value $false
    Write-Host "[OK] Copia en despliegue restaurada con exito: $deployXlsm" -ForegroundColor Green
}

Write-Host "`nRESTAURACION FINALIZADA. El proyecto ha vuelto a la version de referencia 2026-10-17." -ForegroundColor Cyan
