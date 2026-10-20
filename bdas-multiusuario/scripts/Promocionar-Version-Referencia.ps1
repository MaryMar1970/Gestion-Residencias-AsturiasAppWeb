# ==============================================================================
# Script: Promocionar-Version-Referencia.ps1
# Propósito: Guardar la versión actual activa como Versión de Referencia Oficial
#            y registrar copia incremental de cierre en BackupsDiarios.
# ==============================================================================

$ErrorActionPreference = "Stop"

$activeXlsm = "H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm"
$deployXlsm = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"
$backendAccdb = "H:\ResidenciaBD\Residencia_BE.accdb"

$refDir = "H:\ResidenciaBD\Plantilla\VersionReferencia"
$histDir = "H:\ResidenciaBD\Plantilla\VersionReferencia\Historico"
$backupDir = "H:\ResidenciaBD\BackupsDiarios"

$currentRefXlsm = Join-Path $refDir "BDAS_Multiusuario_VERSION_REFERENCIA.xlsm"
$currentRefAccdb = Join-Path $refDir "Residencia_BE_VERSION_REFERENCIA.accdb"

Write-Host "=== 1. Generando Backup Incremental de Cierre de Sesion ===" -ForegroundColor Cyan
$bkpXlsm = Join-Path $backupDir "BDAS_Multiusuario_BACKUP_2026-10-21_0055_VERSION_REFERENCIA.xlsm"
$bkpAccdb = Join-Path $backupDir "Residencia_BE_BACKUP_2026-10-21_0055_VERSION_REFERENCIA.accdb"

Copy-Item -LiteralPath $activeXlsm -Destination $bkpXlsm -Force
Write-Host "[OK] Backup XLSM creado: $bkpXlsm" -ForegroundColor Green

if (Test-Path $backendAccdb) {
    Copy-Item -LiteralPath $backendAccdb -Destination $bkpAccdb -Force
    Write-Host "[OK] Backup ACCDB creado: $bkpAccdb" -ForegroundColor Green
}

Write-Host "=== 2. Archivando Referencia Previa en Historico ===" -ForegroundColor Cyan
$histXlsm = Join-Path $histDir "BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_2315.xlsm"
$histAccdb = Join-Path $histDir "Residencia_BE_VERSION_REFERENCIA_2026-10-20_2315.accdb"

if (Test-Path $currentRefXlsm) {
    Copy-Item -LiteralPath $currentRefXlsm -Destination $histXlsm -Force
    Set-ItemProperty -LiteralPath $histXlsm -Name IsReadOnly -Value $true
    Write-Host "[OK] Historico XLSM archivado: $histXlsm" -ForegroundColor Green
}

if (Test-Path $currentRefAccdb) {
    Copy-Item -LiteralPath $currentRefAccdb -Destination $histAccdb -Force
    Set-ItemProperty -LiteralPath $histAccdb -Name IsReadOnly -Value $true
    Write-Host "[OK] Historico ACCDB archivado: $histAccdb" -ForegroundColor Green
}

Write-Host "=== 3. Promocionando a Nueva Version de Referencia Oficial ===" -ForegroundColor Cyan
if (Test-Path $currentRefXlsm) {
    Set-ItemProperty -LiteralPath $currentRefXlsm -Name IsReadOnly -Value $false
}
Copy-Item -LiteralPath $activeXlsm -Destination $currentRefXlsm -Force
Set-ItemProperty -LiteralPath $currentRefXlsm -Name IsReadOnly -Value $true
Write-Host "[OK] Version de Referencia XLSM actualizada y protegida: $currentRefXlsm" -ForegroundColor Green

if (Test-Path $backendAccdb) {
    if (Test-Path $currentRefAccdb) {
        Set-ItemProperty -LiteralPath $currentRefAccdb -Name IsReadOnly -Value $false
    }
    Copy-Item -LiteralPath $backendAccdb -Destination $currentRefAccdb -Force
    Set-ItemProperty -LiteralPath $currentRefAccdb -Name IsReadOnly -Value $true
    Write-Host "[OK] Version de Referencia ACCDB actualizada y protegida: $currentRefAccdb" -ForegroundColor Green
}

Write-Host "=== 4. Sincronizando Copia de Red H:\ResidenciaBD ===" -ForegroundColor Cyan
Copy-Item -LiteralPath $activeXlsm -Destination $deployXlsm -Force
Set-ItemProperty -LiteralPath $deployXlsm -Name IsReadOnly -Value $false
Write-Host "[OK] Copia desplegada en red sincronizada: $deployXlsm" -ForegroundColor Green
