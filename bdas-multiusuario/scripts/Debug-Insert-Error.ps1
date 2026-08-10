$bdasDir = "H:\ResidenciaApp\bdas-multiusuario"
$ExcelPath = (Get-ChildItem -Path $bdasDir -Filter "*.xlsm" | Select-Object -First 1).FullName
$AccessPath = "H:\ResidenciaBD\Residencia_BE.accdb"

$cn = New-Object -ComObject ADODB.Connection
$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;"
$cn.Open($connStr)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

$wb = $excel.Workbooks.Open($ExcelPath, 0, $true)
$ws = $wb.Sheets.Item("RESIDENCIA OVIEDO")
$maxRow = $ws.Cells($ws.Rows.Count, 1).End(-4162).Row
$arr = $ws.Range($ws.Cells(1, 1), $ws.Cells($maxRow, 30)).Value2
$wb.Close($false)
$excel.Quit()

function EscaparSQL([string]$txt) {
    if ([string]::IsNullOrEmpty($txt)) { return "" }
    return $txt.Replace("'", "''")
}

function FormatearFechaSQL($val) {
    if ($val -is [DateTime]) {
        return "#" + $val.ToString("yyyy-MM-dd HH:mm:ss") + "#"
    }
    if ($val -is [double]) {
        try {
            $dt = [DateTime]::FromOADate($val)
            return "#" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "#"
        } catch {}
    }
    $dt = [DateTime]::Now
    if ([DateTime]::TryParse([string]$val, [ref]$dt)) {
        return "#" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "#"
    }
    return "NULL"
}

$errCount = 0
$dt = [DateTime]::Now
$numOrden = 0
$numFact = 0

for ($r = 2; $r -le $maxRow; $r++) {
    $v = $arr.GetValue($r, 1)
    $numOrden = 0
    if ($v -and [int]::TryParse([string]$v, [ref]$numOrden) -and $numOrden -gt 0) {
        $fechaPet = FormatearFechaSQL ($arr.GetValue($r, 2))
        $numFactVal = $arr.GetValue($r, 3)
        $numFactSql = "NULL"
        $numFact = 0
        if ($numFactVal -and [int]::TryParse([string]$numFactVal, [ref]$numFact) -and $numFact -gt 0) {
            $numFactSql = $numFact
        }
        $dni = EscaparSQL ([string]($arr.GetValue($r, 9)))
        $nombre = EscaparSQL ([string]($arr.GetValue($r, 11)))
        $fechaEnt = FormatearFechaSQL ($arr.GetValue($r, 12))
        $fechaSal = FormatearFechaSQL ($arr.GetValue($r, 13))
        $resolucion = EscaparSQL ([string]($arr.GetValue($r, 16)))
        $numInd = 0; $valInd = [string]($arr.GetValue($r, 17)); [void][int]::TryParse($valInd, [ref]$numInd)
        $numDob = 0; $valDob = [string]($arr.GetValue($r, 18)); [void][int]::TryParse($valDob, [ref]$numDob)
        $habs = EscaparSQL ([string]($arr.GetValue($r, 20)))
        $tel = EscaparSQL ([string]($arr.GetValue($r, 23)))
        $dir = EscaparSQL ([string]($arr.GetValue($r, 24)))
        $cp = EscaparSQL ([string]($arr.GetValue($r, 25)))
        $pob = EscaparSQL ([string]($arr.GetValue($r, 26)))
        $prov = EscaparSQL ([string]($arr.GetValue($r, 27)))
        $pago = EscaparSQL ([string]($arr.GetValue($r, 28)))
        $nowStr = "#" + [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + "#"

        $sql = "INSERT INTO Ordenes (NumOrden, FechaPeticion, NumFactura, Residencia, DNI, Nombre, NumHabIndividuales, NumHabDobles, FechaEntrada, FechaSalida, Resolucion, HabitacionesAsignadas, Telefono, Direccion, CodigoPostal, Poblacion, Provincia, EstadoPago, FechaCreacion, UsuarioCreacion) VALUES ($numOrden, $fechaPet, $numFactSql, 'OVIEDO', '$dni', '$nombre', $numInd, $numDob, $fechaEnt, $fechaSal, '$resolucion', '$habs', '$tel', '$dir', '$cp', '$pob', '$prov', '$pago', $nowStr, 'MIGRACION')"

        try {
            $cn.Execute($sql) | Out-Null
        } catch {
            $errCount++
            if ($errCount -le 10) {
                Write-Host "Fila $r NumOrden: $numOrden EXCEPCION: $_" -ForegroundColor Red
                Write-Host "SQL: $sql" -ForegroundColor Yellow
            }
        }
    }
}

Write-Host "Total errores al insertar RESIDENCIA OVIEDO: $errCount" -ForegroundColor Red
$cn.Close()
