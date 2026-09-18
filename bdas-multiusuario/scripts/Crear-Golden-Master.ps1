$source = 'H:\ResidenciaBD\BDAS_Multiusuario.xlsm'
$goldenDir = 'H:\ResidenciaBD\Plantilla\GoldenMaster'
if (-not (Test-Path $goldenDir)) {
    New-Item -ItemType Directory -Path $goldenDir -Force | Out-Null
}
$goldenFile = Join-Path $goldenDir 'BDAS_Multiusuario_GOLDEN_MASTER.xlsm'
Copy-Item -LiteralPath $source -Destination $goldenFile -Force
(Get-Item $goldenFile).IsReadOnly = $true
Write-Host "[OK] Copia Golden Master blindada en Solo Lectura: $goldenFile"
