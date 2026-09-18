$source = "H:\ResidenciaBD\BDAS_Multiusuario.xlsm"

if (-not (Test-Path $source)) {
    Write-Host "ERROR: No existe $source" -ForegroundColor Red
    exit 1
}

$destFolder = "H:\ResidenciaApp\bdas-multiusuario"
$masterFile = Get-ChildItem "$destFolder\*.xlsm" | Where-Object { $_.Name -notlike "*BACKUP*" -and $_.Name -notlike "*OLD*" } | Select-Object -First 1

if ($masterFile) {
    # 1. Crear copia de seguridad con marca de tiempo del maestro antes de sobrescribirlo
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $backupMaster = Join-Path $destFolder ($masterFile.BaseName + "_PREVIO_SYNC_$timestamp.xlsm")
    Copy-Item -LiteralPath $masterFile.FullName -Destination $backupMaster -Force
    Write-Host "Copia de seguridad creada: $backupMaster" -ForegroundColor Yellow

    # 2. Promover BDAS_Multiusuario.xlsm como nuevo Maestro en ResidenciaApp
    Copy-Item -LiteralPath $source -Destination $masterFile.FullName -Force
    Write-Host "BDAS_Multiusuario.xlsm promovido como MAESTRO en: $($masterFile.FullName)" -ForegroundColor Green
} else {
    Write-Host "ERROR: No se localizo el archivo maestro en $destFolder" -ForegroundColor Red
}
