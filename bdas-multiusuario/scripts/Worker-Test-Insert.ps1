param(
    [int]$Id,
    [string]$AccessPath = "H:\ResidenciaBD\Residencia_BE.accdb"
)

$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;"
$retries = 10
$success = $false

while ($retries -gt 0 -and -not $success) {
    try {
        $conn = New-Object -ComObject ADODB.Connection
        $conn.Open($connStr)
        
        $conn.BeginTrans()
        
        $rs = New-Object -ComObject ADODB.Recordset
        $rs.Open("SELECT MAX(NumOrden) AS MaxOrden FROM Ordenes WHERE Residencia = 'GIJON'", $conn)
        $nextOrden = 1
        if (-not $rs.EOF -and -not [DBNull]::Value.Equals($rs.Fields.Item("MaxOrden").Value)) {
            $nextOrden = [int]$rs.Fields.Item("MaxOrden").Value + 1
        }
        $rs.Close()
        
        $dni = "TEST_CONC_$Id"
        $nombre = "TESTER CONCURRENTE $Id"
        $cmd = New-Object -ComObject ADODB.Command
        $cmd.ActiveConnection = $conn
        $cmd.CommandText = "INSERT INTO Ordenes (NumOrden, Residencia, FechaPeticion, DNI, Nombre, Apellidos, TipoHuesped, Resolucion, EstadoPago, FechaCreacion) VALUES ($nextOrden, 'GIJON', Date(), '$dni', '$nombre', 'PRUEBA', 'ESTUDIANTE', 'Pendiente', 'Pendiente', Date())"
        $cmd.Execute()
        
        $conn.CommitTrans()
        $conn.Close()
        $success = $true
        Write-Host "OK ${Id} ${nextOrden}"
    } catch {
        if ($conn -and $conn.State -eq 1) {
            try { $conn.RollbackTrans() } catch {}
            try { $conn.Close() } catch {}
        }
        $retries--
        Start-Sleep -Milliseconds (Get-Random -Minimum 50 -Maximum 150)
    }
}
