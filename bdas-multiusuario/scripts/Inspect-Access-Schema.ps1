$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\ResidenciaBD\Residencia_BE.accdb;"
$cn = New-Object -ComObject ADODB.Connection
$cn.Open($connStr)
$rs = New-Object -ComObject ADODB.Recordset
$rs.Open("SELECT TOP 1 * FROM Ordenes", $cn, 1, 1)

Write-Host "=== FIELDS IN TABLE 'Ordenes' IN ACCESS ==="
for ($i = 0; $i -lt $rs.Fields.Count; $i++) {
    $f = $rs.Fields.Item($i)
    Write-Host "Field [$i]: $($f.Name) (Type: $($f.Type))"
}

$rs.Close()
$cn.Close()
