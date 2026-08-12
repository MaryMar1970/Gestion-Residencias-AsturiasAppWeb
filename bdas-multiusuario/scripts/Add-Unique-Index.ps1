$dbPath = "H:\ResidenciaBD\Residencia_BE.accdb"
$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$dbPath;"
$cn = New-Object -ComObject ADODB.Connection

try {
    $cn.Open($connStr)
    $cn.Execute("CREATE UNIQUE INDEX idx_Residencia_NumOrden ON Ordenes (Residencia, NumOrden);")
    Write-Host "[OK] Indice unico idx_Residencia_NumOrden creado." -ForegroundColor Green
} catch {
    Write-Host "[INFO] Indice ya existente o aviso: $_" -ForegroundColor Yellow
} finally {
    if ($cn.State -eq 1) { $cn.Close() }
}
