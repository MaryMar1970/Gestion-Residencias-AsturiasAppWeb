# ==============================================================================
# Script: Restaurar-Golden-Master.ps1
# Propósito: Recuperar en 1 segundo el archivo maestro con textos y acentos perfectos
#            en caso de que se haya modificado, corrompido o guardado por error.
# ==============================================================================

param(
    [switch]$Force
)

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "      RESTAURACION DE VERSION MAESTRA DE TEXTOS (GOLDEN MASTER)  " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

$goldenDir = "H:\ResidenciaBD\Plantilla\GoldenMaster"
$goldenFile = Join-Path $goldenDir "BDAS_Multiusuario_GOLDEN_MASTER.xlsm"
$deployFile = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"
$repoFolder = "H:\ResidenciaApp\bdas-multiusuario"
$repoFile = Get-ChildItem "$repoFolder\*.xlsm" | Where-Object { $_.Name -notlike "*BACKUP*" -and $_.Name -notlike "*OLD*" -and $_.Name -notlike "*PREVIO*" } | Select-Object -First 1

if (-not (Test-Path $goldenFile)) {
    Write-Host "ERROR: No se encuentra la copia protegida en $goldenFile" -ForegroundColor Red
    exit 1
}

# 1. Crear backup de seguridad del estado actual antes de sobreescribir (por si se quiere revisar algo)
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
if (Test-Path $deployFile) {
    $bkp = "H:\ResidenciaBD\BDAS_Multiusuario_ERRONEO_$timestamp.xlsm"
    Copy-Item -LiteralPath $deployFile -Destination $bkp -Force
    Write-Host "[INFO] Guardada copia del archivo actual en: $bkp" -ForegroundColor Gray
}

# 2. Restaurar en H:\ResidenciaBD\BDAS_Multiusuario.xlsm
Copy-Item -LiteralPath $goldenFile -Destination $deployFile -Force
Write-Host "[OK] H:\ResidenciaBD\BDAS_Multiusuario.xlsm restaurado con exito al Golden Master." -ForegroundColor Green

# 3. Restaurar en el repositorio maestro si existe
if ($repoFile) {
    Copy-Item -LiteralPath $goldenFile -Destination $repoFile.FullName -Force
    Write-Host "[OK] $($repoFile.FullName) restaurado con exito al Golden Master." -ForegroundColor Green
}

Write-Host "`nRESTAURACION FINALIZADA. El archivo vuelve a tener todos tus acentos y textos limpios." -ForegroundColor Cyan
