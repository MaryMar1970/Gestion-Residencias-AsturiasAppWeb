# Script de Pruebas de Concurrencia — BDAS Multiusuario
# Simula peticiones concurrentes simultáneas sobre H:\ResidenciaBD\Residencia_BE.accdb

param(
    [string]$AccessPath = "H:\ResidenciaBD\Residencia_BE.accdb",
    [int]$NumThreads = 10
)

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "  PRUEBA DE CONCURRENCIA MULTIUSUARIO - BDAS ACCESS BACK-END" -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "Ruta DB: $AccessPath"
Write-Host "Hilos concurrentes a simular: $NumThreads"
Write-Host ""

if (-not (Test-Path $AccessPath)) {
    Write-Host "ERROR: No se encuentra la base de datos en $AccessPath" -ForegroundColor Red
    exit 1
}

$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;"

# 1. Obtener el máximo NumOrden actual
$conn = New-Object -ComObject ADODB.Connection
$conn.Open($connStr)
$rs = New-Object -ComObject ADODB.Recordset
$rs.Open("SELECT MAX(NumOrden) AS MaxOrden FROM Ordenes", $conn)
$initialMaxOrden = 0
if (-not $rs.EOF -and -not [DBNull]::Value.Equals($rs.Fields.Item("MaxOrden").Value)) {
    $initialMaxOrden = [int]$rs.Fields.Item("MaxOrden").Value
}
$rs.Close()
$conn.Close()

Write-Host "Nº Orden actual máximo en la BD: $initialMaxOrden" -ForegroundColor Yellow

# 2. Ejecutar inserciones concurrentes en paralelo llamando a Worker-Test-Insert.ps1 en 32-bit
Write-Host "`nLanzando $NumThreads inserciones atómicas en paralelo..." -ForegroundColor Cyan

$ps32 = "C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe"
$workerScript = Join-Path $PSScriptRoot "Worker-Test-Insert.ps1"
$jobs = @()

for ($i = 1; $i -le $NumThreads; $i++) {
    $jobs += Start-Process -FilePath $ps32 -ArgumentList "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", "`"$workerScript`"", "-Id", $i, "-AccessPath", "`"$AccessPath`"" -NoNewWindow -PassThru
}

$jobs | Wait-Process

# Recopilar resultados de la BD
$connCheck = New-Object -ComObject ADODB.Connection
$connCheck.Open($connStr)
$rsCheck = New-Object -ComObject ADODB.Recordset
$rsCheck.Open("SELECT NumOrden, DNI FROM Ordenes WHERE DNI LIKE 'TEST_CONC_%' ORDER BY NumOrden", $connCheck)
$results = @()
while (-not $rsCheck.EOF) {
    $results += [PSCustomObject]@{
        DNI = $rsCheck.Fields.Item("DNI").Value
        NumOrden = $rsCheck.Fields.Item("NumOrden").Value
    }
    $rsCheck.MoveNext()
}
$rsCheck.Close()
$connCheck.Close()

Write-Host "`nResultados de las inserciones guardadas en BD:" -ForegroundColor Cyan
$results | Format-Table -AutoSize

# 3. Verificación de integridad y no duplicación de NumOrden
$insertedOrders = $results.NumOrden
$uniqueOrders = $insertedOrders | Select-Object -Unique

Write-Host "`n--- RESUMEN DE INTEGRIDAD ---" -ForegroundColor Yellow
Write-Host "Inserciones exitosas: $($results.Count) / $NumThreads"
Write-Host "Nº Orden unicos generados: $($uniqueOrders.Count)"

if (($results.Count -eq $NumThreads) -and ($uniqueOrders.Count -eq $NumThreads)) {
    Write-Host "[OK] PRUEBA DE CONCURRENCIA SUPERADA AL 100%: Sin colisiones de NumOrden ni bloqueos." -ForegroundColor Green
} else {
    Write-Host "[ERROR] ATENCION: Se detectaron colisiones o fallos en inserciones concurrentes." -ForegroundColor Red
}

# 4. Limpieza de registros de prueba
Write-Host "`nLimpiando registros de prueba de la BD..." -ForegroundColor Gray
$connClean = New-Object -ComObject ADODB.Connection
$connClean.Open($connStr)
$cmdClean = New-Object -ComObject ADODB.Command
$cmdClean.ActiveConnection = $connClean
$cmdClean.CommandText = "DELETE FROM Ordenes WHERE DNI LIKE 'TEST_CONC_%'"
$cmdClean.Execute()
$connClean.Close()

Write-Host "Limpieza completada." -ForegroundColor Gray
