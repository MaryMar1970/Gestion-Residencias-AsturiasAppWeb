$timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$sourceDeploy = 'H:\ResidenciaBD\BDAS_Multiusuario.xlsm'
$backupDir = 'H:\ResidenciaBD\BackupsDiarios'
if (-not (Test-Path $backupDir)) { 
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null 
}
$destBkp = Join-Path $backupDir "BDAS_Multiusuario_BACKUP_$timestamp.xlsm"
Copy-Item -LiteralPath $sourceDeploy -Destination $destBkp -Force
Write-Host "[OK] Copia de seguridad diaria de Excel generada con exito en: $destBkp"

$sourceAcc = 'H:\ResidenciaBD\Residencia_BE.accdb'
if (Test-Path $sourceAcc) {
    $destBkpAcc = Join-Path $backupDir "Residencia_BE_BACKUP_$timestamp.accdb"
    Copy-Item -LiteralPath $sourceAcc -Destination $destBkpAcc -Force
    Write-Host "[OK] Copia de seguridad diaria de Access generada con exito en: $destBkpAcc"
}
