# Script de Despliegue — BDAS Multiusuario Front-End
# Copia el libro Excel maestro a H:\ResidenciaBD\BDAS_Multiusuario.xlsm y crea el acceso directo en el escritorio.

param(
    [string]$SourceExcel = "",
    [string]$TargetDir   = "H:\ResidenciaBD",
    [string]$TargetName  = "BDAS_Multiusuario.xlsm"
)

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "      DESPLIEGUE FINAL - BDAS MULTIUSUARIO FRONT-END" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

if ([string]::IsNullOrEmpty($SourceExcel)) {
    $found = Get-ChildItem "H:\ResidenciaApp\bdas-multiusuario\*.xlsm" | Select-Object -First 1
    if ($found) { $SourceExcel = $found.FullName }
}

if (-not (Test-Path $SourceExcel)) {
    Write-Host "ERROR: No se encuentra el archivo maestro en $SourceExcel" -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $TargetDir)) {
    New-Item -ItemType Directory -Path $TargetDir -Force | Out-Null
    Write-Host "Directorio $TargetDir creado con exito." -ForegroundColor Green
}

$targetPath = Join-Path $TargetDir $TargetName

Write-Host "Copiando libro Excel maestro a la ubicacion final compartida..." -ForegroundColor Yellow
Write-Host "Origen : $SourceExcel"
Write-Host "Destino: $targetPath"

Copy-Item -LiteralPath $SourceExcel -Destination $targetPath -Force

if (Test-Path $targetPath) {
    Write-Host "`n[OK] Libro Excel Front-End copiado correctamente a $targetPath" -ForegroundColor Green
} else {
    Write-Host "`n[ERROR] Error al copiar el libro Excel Front-End." -ForegroundColor Red
    exit 1
}

# Crear acceso directo en el Escritorio del usuario actual
$desktopPath = [System.Environment]::GetFolderPath([System.Environment+SpecialFolder]::Desktop)
$shortcutPath = Join-Path $desktopPath "BDAS Multiusuario.lnk"

$wshShell = New-Object -ComObject WScript.Shell
$shortcut = $wshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $targetPath
$shortcut.Description = "BDAS Multiusuario (Excel + Access)"
$shortcut.IconLocation = "excel.exe,0"
$shortcut.Save()

Write-Host "[OK] Acceso directo creado en el escritorio: $shortcutPath" -ForegroundColor Green
Write-Host "`nDESPLIEGUE COMPLETADO CON EXITO." -ForegroundColor Cyan
