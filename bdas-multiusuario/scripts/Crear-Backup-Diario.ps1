$timestamp = Get-Date -Format 'yyyy-MM-dd_HHmm'
$sourceDeploy = 'H:\ResidenciaBD\BDAS_Multiusuario.xlsm'
$backupDir = 'H:\ResidenciaBD\BackupsDiarios'
if (-not (Test-Path $backupDir)) { 
    New-Item -ItemType Directory -Path $backupDir -Force | Out-Null 
}
$destBkp = Join-Path $backupDir "BDAS_Multiusuario_BACKUP_$timestamp.xlsm"
Copy-Item -LiteralPath $sourceDeploy -Destination $destBkp -Force
Write-Host "[OK] Copia de seguridad diaria generada con exito en: $destBkp"
