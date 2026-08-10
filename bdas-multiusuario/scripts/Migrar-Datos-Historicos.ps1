$bdasDir = "H:\ResidenciaApp\bdas-multiusuario"
$ExcelPath = (Get-ChildItem -Path $bdasDir -Filter "*.xlsm" | Select-Object -First 1).FullName
$AccessPath = "H:\ResidenciaBD\Residencia_BE.accdb"

Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "  BDAS - Migracion de Datos Historicos a Access (Fase 3)       " -ForegroundColor Cyan
Write-Host "================================================================" -ForegroundColor Cyan
Write-Host "Excel Origen: $ExcelPath" -ForegroundColor Yellow

$cn = New-Object -ComObject ADODB.Connection
$connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=$AccessPath;"
$cn.Open($connStr)
Write-Host "[OK] Conectado a Access." -ForegroundColor Green

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.EnableEvents = $false

Write-Host "[PROCESANDO] Abriendo libro Excel..." -ForegroundColor Cyan
$wb = $excel.Workbooks.Open($ExcelPath, 0, $true)

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

$cn.BeginTrans()

try {
    $residencias = @(
        @{ Pattern="RESIDENCIA*GIJ*"; Code="GIJON"; ColHab=19; ColTel=22; ColDir=23; ColCP=24; ColPob=25; ColProv=26; ColPago=27 },
        @{ Pattern="RESIDENCIA*SOTO*"; Code="SOTO"; ColHab=19; ColTel=22; ColDir=23; ColCP=24; ColPob=25; ColProv=26; ColPago=27 },
        @{ Pattern="RESIDENCIA*OVIEDO*"; Code="OVIEDO"; ColHab=20; ColTel=23; ColDir=24; ColCP=25; ColPob=26; ColProv=27; ColPago=28 }
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

        if (-not $targetWs) { continue }

        $sheetName = $targetWs.Name
        $code = $res.Code
        $maxRow = $targetWs.Cells($targetWs.Rows.Count, 1).End(-4162).Row
        if ($maxRow -lt 2) { continue }

        # Carga ultra-rapida en memoria (1 sola llamada COM por hoja)
        $range = $targetWs.Range($targetWs.Cells(1, 1), $targetWs.Cells($maxRow, 30))
        $arr = $range.Value2

        $count = 0
        $numOrden = 0
        $numFact = 0

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

                $dni = EscaparSQL ([string]($arr.GetValue($r, 9)))
                $nombre = EscaparSQL ([string]($arr.GetValue($r, 11)))
                $fechaEnt = FormatearFechaSQL ($arr.GetValue($r, 12))
                $fechaSal = FormatearFechaSQL ($arr.GetValue($r, 13))
                $resolucion = EscaparSQL ([string]($arr.GetValue($r, 16)))

                $numInd = 0; $valInd = [string]($arr.GetValue($r, 17)); [void][int]::TryParse($valInd, [ref]$numInd)
                $numDob = 0; $valDob = [string]($arr.GetValue($r, 18)); [void][int]::TryParse($valDob, [ref]$numDob)

                $habs = EscaparSQL ([string]($arr.GetValue($r, $res.ColHab)))
                $tel = EscaparSQL ([string]($arr.GetValue($r, $res.ColTel)))
                $dir = EscaparSQL ([string]($arr.GetValue($r, $res.ColDir)))
                $cp = EscaparSQL ([string]($arr.GetValue($r, $res.ColCP)))
                $pob = EscaparSQL ([string]($arr.GetValue($r, $res.ColPob)))
                $prov = EscaparSQL ([string]($arr.GetValue($r, $res.ColProv)))
                $pago = EscaparSQL ([string]($arr.GetValue($r, $res.ColPago)))

                $nowStr = "#" + [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + "#"

                $sql = "INSERT INTO Ordenes (NumOrden, FechaPeticion, NumFactura, Residencia, DNI, Nombre, NumHabIndividuales, NumHabDobles, FechaEntrada, FechaSalida, Resolucion, HabitacionesAsignadas, Telefono, Direccion, CodigoPostal, Poblacion, Provincia, EstadoPago, FechaCreacion, UsuarioCreacion) VALUES ($numOrden, $fechaPet, $numFactSql, '$code', '$dni', '$nombre', $numInd, $numDob, $fechaEnt, $fechaSal, '$resolucion', '$habs', '$tel', '$dir', '$cp', '$pob', '$prov', '$pago', $nowStr, 'MIGRACION')"

                try {
                    $cn.Execute($sql) | Out-Null
                    $count++
                } catch {}
            }
        }
        Write-Host "[OK] $sheetName -> $count ordenes migradas." -ForegroundColor Green
        $totalOrdenes += $count
    }

    # --- MIGRAR LISTA NEGRA ---
    $wsLN = $null
    foreach ($s in $wb.Sheets) { if ($s.Name -like "*LISTA*NEGRA*") { $wsLN = $s; break } }
    $countLN = 0
    if ($wsLN) {
        $maxLN = $wsLN.Cells($wsLN.Rows.Count, 1).End(-4162).Row
        if ($maxLN -ge 2) {
            $arrLN = $wsLN.Range($wsLN.Cells(1, 1), $wsLN.Cells($maxLN, 5)).Value2
            for ($r = 2; $r -le $maxLN; $r++) {
                $dniLN = EscaparSQL ([string]($arrLN.GetValue($r, 1)))
                if ($dniLN.Length -gt 0) {
                    $nomLN = EscaparSQL ([string]($arrLN.GetValue($r, 2)))
                    $motLN = EscaparSQL ([string]($arrLN.GetValue($r, 3)))
                    $nowStr = "#" + [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + "#"
                    $sqlLN = "INSERT INTO ListaNegra (DNI, Nombre, Motivo, FechaAlta, Activo) VALUES ('$dniLN', '$nomLN', '$motLN', $nowStr, True)"
                    try { $cn.Execute($sqlLN) | Out-Null; $countLN++ } catch {}
                }
            }
        }
    }
    Write-Host "[OK] LISTA NEGRA -> $countLN registros migrados." -ForegroundColor Green

    # --- MIGRAR LOGS ---
    $countLog = 0
    foreach ($s in $wb.Sheets) {
        if ($s.Name -like "*LOG*") {
            $maxLog = $s.Cells($s.Rows.Count, 1).End(-4162).Row
            if ($maxLog -ge 2) {
                $arrLog = $s.Range($s.Cells(1, 1), $s.Cells($maxLog, 5)).Value2
                for ($r = 2; $r -le $maxLog; $r++) {
                    $userLog = EscaparSQL ([string]($arrLog.GetValue($r, 2)))
                    if ($userLog.Length -gt 0) {
                        $accLog = EscaparSQL ([string]($arrLog.GetValue($r, 3)))
                        $detLog = EscaparSQL ([string]($arrLog.GetValue($r, 4)))
                        $nowStr = "#" + [DateTime]::Now.ToString("yyyy-MM-dd HH:mm:ss") + "#"
                        $sqlLog = "INSERT INTO LogActividad (FechaHora, Usuario, Accion, Detalle) VALUES ($nowStr, '$userLog', '$accLog', '$detLog')"
                        try { $cn.Execute($sqlLog) | Out-Null; $countLog++ } catch {}
                    }
                }
            }
        }
    }
    Write-Host "[OK] LOGS -> $countLog registros migrados." -ForegroundColor Green

    $cn.CommitTrans()
    Write-Host ""
    Write-Host "================================================================" -ForegroundColor Cyan
    Write-Host "  MIGRACION COMPLETADA EXITOSAMENTE                             " -ForegroundColor Green
    Write-Host "  • Total Ordenes: $totalOrdenes" -ForegroundColor Yellow
    Write-Host "  • Total Lista Negra: $countLN" -ForegroundColor Yellow
    Write-Host "  • Total Log Actividad: $countLog" -ForegroundColor Yellow
    Write-Host "================================================================" -ForegroundColor Cyan

} catch {
    Write-Error "Error durante la migracion: $_"
    try { $cn.RollbackTrans() } catch {}
} finally {
    if ($wb) { $wb.Close($false) }
    if ($excel) { $excel.Quit() }
    if ($cn) { $cn.Close() }
}
