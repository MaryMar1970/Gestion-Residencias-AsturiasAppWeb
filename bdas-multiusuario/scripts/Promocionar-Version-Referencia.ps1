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

Write-Host "=== 1. Generando Backup Incremental de Cierre de Sesion (Hito 2026-10-22 19:45) ===" -ForegroundColor Cyan
$bkpXlsm = Join-Path $backupDir "BDAS_Multiusuario_BACKUP_2026-10-22_1945_VERSION_REFERENCIA.xlsm"
$bkpAccdb = Join-Path $backupDir "Residencia_BE_BACKUP_2026-10-22_1945_VERSION_REFERENCIA.accdb"

Copy-Item -LiteralPath $activeXlsm -Destination $bkpXlsm -Force
Write-Host "[OK] Backup XLSM creado: $bkpXlsm" -ForegroundColor Green

if (Test-Path $backendAccdb) {
    Copy-Item -LiteralPath $backendAccdb -Destination $bkpAccdb -Force
    Write-Host "[OK] Backup ACCDB creado: $bkpAccdb" -ForegroundColor Green
}

Write-Host "=== 2. Archivando Referencias Previas en Historico ===" -ForegroundColor Cyan
$hist1855Xlsm = Join-Path $histDir "BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-22_1855.xlsm"
$hist1855Accdb = Join-Path $histDir "Residencia_BE_VERSION_REFERENCIA_2026-10-22_1855.accdb"

$prev1855Xlsm = Join-Path $backupDir "BDAS_Multiusuario_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.xlsm"
$prev1855Accdb = Join-Path $backupDir "Residencia_BE_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.accdb"

if ((-not (Test-Path $hist1855Xlsm)) -and (Test-Path $prev1855Xlsm)) {
    Copy-Item -LiteralPath $prev1855Xlsm -Destination $hist1855Xlsm -Force
    Set-ItemProperty -LiteralPath $hist1855Xlsm -Name IsReadOnly -Value $true
    Write-Host "[OK] Historico 18:55 XLSM archivado: $hist1855Xlsm" -ForegroundColor Green
}

if ((-not (Test-Path $hist1855Accdb)) -and (Test-Path $prev1855Accdb)) {
    Copy-Item -LiteralPath $prev1855Accdb -Destination $hist1855Accdb -Force
    Set-ItemProperty -LiteralPath $hist1855Accdb -Name IsReadOnly -Value $true
    Write-Host "[OK] Historico 18:55 ACCDB archivado: $hist1855Accdb" -ForegroundColor Green
}

Write-Host "=== 3. Promocionando a Nueva Version de Referencia Oficial (Hito 2026-10-22 19:45) ===" -ForegroundColor Cyan
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
