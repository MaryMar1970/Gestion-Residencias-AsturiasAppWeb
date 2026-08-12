$files = Get-ChildItem "H:\ResidenciaApp\bdas-multiusuario\scripts\*" -Include "*.frm", "*.bas", "*.vbs"

foreach ($file in $files) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    $nonAsciiCount = 0
    for ($i = 0; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i] -gt 127) {
            $nonAsciiCount++
        }
    }
    Write-Host "File: $($file.Name) -> Non-ASCII bytes: $nonAsciiCount"
}
