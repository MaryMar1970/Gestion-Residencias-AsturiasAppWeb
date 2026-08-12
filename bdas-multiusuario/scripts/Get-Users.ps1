$conn = New-Object -ComObject ADODB.Connection
$conn.Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\ResidenciaBD\Residencia_BE.accdb")
$rs = $conn.Execute("SELECT NombreUsuario, NombreCompleto, Rol, Residencias, Activo, Clave FROM Usuarios")

while (-not $rs.EOF) {
    $u = $rs.Fields.Item("NombreUsuario").Value
    $n = $rs.Fields.Item("NombreCompleto").Value
    $r = $rs.Fields.Item("Rol").Value
    $res = $rs.Fields.Item("Residencias").Value
    $a = $rs.Fields.Item("Activo").Value
    $c = $rs.Fields.Item("Clave").Value
    Write-Host "User: '$u' | Name: '$n' | Role: '$r' | Residencias: '$res' | Active: $a | PassHash: '$c'"
    $rs.MoveNext()
}
$conn.Close()
