param(
    [string]$SourceExcel = "",
    [string]$AccessPath  = "H:\ResidenciaBD\Residencia_BE.accdb"
)

Write-Host "=================================================================" -ForegroundColor Cyan
Write-Host "      MIGRACION EXHAUSTIVA DE DATOS HISTORICOS A ACCESS         " -ForegroundColor Cyan
Write-Host "=================================================================" -ForegroundColor Cyan

if ([string]::IsNullOrEmpty($SourceExcel)) {
    $found = Get-ChildItem "H:\ResidenciaApp\bdas-multiusuario\*_OLD.xlsm" | Select-Object -First 1
    if (-not $found) {
        $found = Get-ChildItem "H:\ResidenciaApp\bdas-multiusuario\*.xlsm" | Where-Object { $_.Name -notlike "*BACKUP*" } | Select-Object -First 1
    }
    if ($found) { $SourceExcel = $found.FullName }
}

if (-not (Test-Path $SourceExcel)) {
    Write-Host "ERROR: No se encuentra el archivo maestro en $SourceExcel" -ForegroundColor Red
    exit 1
}

Write-Host "Origen : $SourceExcel"
Write-Host "Destino: $AccessPath"

$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;"
$cn = New-Object -ComObject ADODB.Connection
$cn.Open($connStr)

function EscaparSQL([string]$t) {
    if ([string]::IsNullOrEmpty($t)) { return "" }
    return $t.Replace("'", "''").Trim()
}

function FormatearFechaSQL($val) {
    if (-not $val) { return "NULL" }
    $dt = [DateTime]::MinValue
    if ([DateTime]::TryParse([string]$val, [ref]$dt)) {
        return "#" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "#"
    }
    if ($val -is [double] -or $val -is [int]) {
        $dt = [DateTime]::FromOADate([double]$val)
        return "#" + $dt.ToString("yyyy-MM-dd HH:mm:ss") + "#"
    }
    return "NULL"
}

function ValorNumEntero($val) {
    if (-not $val) { return "NULL" }
    $n = 0
    if ([int]::TryParse([string]$val, [ref]$n)) { return $n }
    if ($val -is [double]) { return [int]$val }
    return "NULL"
}

function ValorNumDouble($val) {
    if (-not $val) { return "NULL" }
    $d = 0.0
    $s = [string]$val
    $s = $s.Replace(",", ".")
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) {
        return $d.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    }
    return "NULL"
}

$cn.BeginTrans()

try {
    # 1. Limpiar tabla Ordenes para re-migración limpia
    Write-Host "Limpiando tabla Ordenes en Access..." -ForegroundColor Yellow
    $cn.Execute("DELETE FROM Ordenes;") | Out-Null

    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false
    $xl.DisplayAlerts = $false
    $xl.EnableEvents = $false

    $wb = $xl.Workbooks.Open($SourceExcel, 0, $true)

    $residencias = @(
        @{ Pattern="RESIDENCIA*GIJ*";   Code="GIJON";  IsOviedo=$false; ColHab=19; ColDto=20; ColImp=21; ColTel=22; ColDir=23; ColCP=24; ColPob=25; ColProv=26; ColPago=27; ColGrab=28 },
        @{ Pattern="RESIDENCIA*SOTO*";  Code="SOTO";   IsOviedo=$false; ColHab=19; ColDto=20; ColImp=21; ColTel=22; ColDir=23; ColCP=24; ColPob=25; ColProv=26; ColPago=27; ColGrab=28 },
        @{ Pattern="RESIDENCIA*OVIEDO*"; Code="OVIEDO"; IsOviedo=$true;  ColSuple=19; ColHab=20; ColDto=21; ColImp=22; ColTel=23; ColDir=24; ColCP=25; ColPob=26; ColProv=27; ColPago=28; ColGrab=32 }
    )

    $totalOrdenes = 0

    foreach ($res in $residencias) {
        $targetWs = $null
        foreach ($s in $wb.Sheets) {
            if ($s.Name -like $res.Pattern) {
                $targetWs = $s
                break
            }
        }

        if (-not $targetWs) {
            Write-Host "AVISO: No se encontro hoja para $($res.Code)" -ForegroundColor Yellow
            continue
        }

        $sheetName = $targetWs.Name
        $code = $res.Code
        $maxRow = $targetWs.Cells($targetWs.Rows.Count, 1).End(-4162).Row
        if ($maxRow -lt 2) { continue }

        $range = $targetWs.Range($targetWs.Cells(1, 1), $targetWs.Cells($maxRow, 32))
        $arr = $range.Value2

        $count = 0

        for ($r = 2; $r -le $maxRow; $r++) {
            $numOrdenVal = $arr.GetValue($r, 1)
            $numOrden = 0
            if ($numOrdenVal -and [int]::TryParse([string]$numOrdenVal, [ref]$numOrden) -and $numOrden -gt 0) {
                
                $fechaPet = FormatearFechaSQL ($arr.GetValue($r, 2))
                
                $numFactVal = $arr.GetValue($r, 3)
                $numFactSql = "NULL"
                $numFact = 0
                if ($numFactVal -and [int]::TryParse([string]$numFactVal, [ref]$numFact) -and $numFact -gt 0) {
                    $numFactSql = $numFact
                }

                $finalidad  = EscaparSQL ([string]($arr.GetValue($r, 4)))
                $empleo     = EscaparSQL ([string]($arr.GetValue($r, 5)))
                $situacion  = EscaparSQL ([string]($arr.GetValue($r, 6)))
                $evaluacion = EscaparSQL ([string]($arr.GetValue($r, 7)))
                
                $col8Val    = EscaparSQL ([string]($arr.GetValue($r, 8)))
                $comision   = ""
                $turno      = ""
                if ($code -eq "SOTO") {
                    $turno = $col8Val
                } else {
                    $comision = $col8Val
                }

                $dni        = EscaparSQL ([string]($arr.GetValue($r, 9)))
                $rango      = EscaparSQL ([string]($arr.GetValue($r, 10)))
                $nombre     = EscaparSQL ([string]($arr.GetValue($r, 11)))
                $fechaEnt   = FormatearFechaSQL ($arr.GetValue($r, 12))
                $fechaSal   = FormatearFechaSQL ($arr.GetValue($r, 13))
                $diasUso    = ValorNumEntero ($arr.GetValue($r, 14))
                $pax        = ValorNumEntero ($arr.GetValue($r, 15))
                $resolucion = EscaparSQL ([string]($arr.GetValue($r, 16)))

                $numInd     = ValorNumEntero ($arr.GetValue($r, 17))
                $numDob     = ValorNumEntero ($arr.GetValue($r, 18))

                $camaSuple  = ""
                if ($res.IsOviedo) {
                    $camaSuple = EscaparSQL ([string]($arr.GetValue($r, $res.ColSuple)))
                }

                $habs       = EscaparSQL ([string]($arr.GetValue($r, $res.ColHab)))
                $dtoFam     = EscaparSQL ([string]($arr.GetValue($r, $res.ColDto)))
                $importe    = ValorNumDouble ($arr.GetValue($r, $res.ColImp))
                $tel        = EscaparSQL ([string]($arr.GetValue($r, $res.ColTel)))
                $dir        = EscaparSQL ([string]($arr.GetValue($r, $res.ColDir)))
                $cp         = EscaparSQL ([string]($arr.GetValue($r, $res.ColCP)))
                $pob        = EscaparSQL ([string]($arr.GetValue($r, $res.ColPob)))
                $prov       = EscaparSQL ([string]($arr.GetValue($r, $res.ColProv)))
                $pago       = EscaparSQL ([string]($arr.GetValue($r, $res.ColPago)))
                $grab       = EscaparSQL ([string]($arr.GetValue($r, $res.ColGrab)))

                $nowStr = "#" + [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + "#"

                $sql = "INSERT INTO Ordenes (" +
                       "NumOrden, FechaPeticion, NumFactura, Residencia, Finalidad, Empleo, Situacion, Evaluacion, " +
                       "Comision, Turno, DNI, Rango, Nombre, FechaEntrada, FechaSalida, DiasUso, PAX, Resolucion, " +
                       "NumHabIndividuales, NumHabDobles, CamaSuple, HabitacionesAsignadas, DtoFamNum, Importe, " +
                       "Telefono, Direccion, CodigoPostal, Poblacion, Provincia, EstadoPago, FechaGrabacion, " +
                       "FechaCreacion, UsuarioCreacion" +
                       ") VALUES (" +
                       "$numOrden, $fechaPet, $numFactSql, '$code', '$finalidad', '$empleo', '$situacion', '$evaluacion', " +
                       "'$comision', '$turno', '$dni', '$rango', '$nombre', $fechaEnt, $fechaSal, $diasUso, $pax, '$resolucion', " +
                       "$numInd, $numDob, '$camaSuple', '$habs', '$dtoFam', $importe, " +
                       "'$tel', '$dir', '$cp', '$pob', '$prov', '$pago', '$grab', " +
                       "$nowStr, 'MIGRACION')"

                try {
                    $cn.Execute($sql) | Out-Null
                    $count++
                } catch {
                    Write-Host "Error en fila $r ($numOrden): $_" -ForegroundColor Red
                }
            }
        }
        Write-Host "[OK] $sheetName -> $count ordenes migradas exhaustivamente." -ForegroundColor Green
        $totalOrdenes += $count
    }

    $wb.Close($false)
    $xl.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null

    $cn.CommitTrans()
    Write-Host "`n[EXITO] TOTAL ORDENES INTEGRADAS EN ACCESS: $totalOrdenes" -ForegroundColor Green

} catch {
    Write-Host "ERROR CRITICO: $_" -ForegroundColor Red
    if ($cn -and $cn.State -eq 1) {
        $cn.RollbackTrans()
    }
} finally {
    if ($cn -and $cn.State -eq 1) {
        $cn.Close()
    }
}
