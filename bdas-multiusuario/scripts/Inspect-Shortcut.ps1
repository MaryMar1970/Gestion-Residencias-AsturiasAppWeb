$desktop = [System.Environment]::GetFolderPath("Desktop")
$lnkFile = Join-Path $desktop "BDAS Multiusuario.lnk"

if (Test-Path $lnkFile) {
    $sh = New-Object -ComObject WScript.Shell
    $target = $sh.CreateShortcut($lnkFile).TargetPath
    Write-Host "Shortcut target: $target"
} else {
    Write-Host "Shortcut not found on desktop: $lnkFile"
    Get-ChildItem $desktop -Filter "*.lnk" | ForEach-Object { Write-Host "Found LNK: $($_.FullName)" }
}
