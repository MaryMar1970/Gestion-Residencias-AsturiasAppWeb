$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $xl.Workbooks.Open('H:\ResidenciaBD\BDAS_Multiusuario.xlsm', 0, $true)

function DumpHeaders($sheetName) {
    Write-Host "=== HEADERS FOR: [$sheetName] ==="
    try {
        $ws = $wb.Sheets.Item($sheetName)
        for ($col = 1; $col -le 35; $col++) {
            $val = $ws.Cells.Item(1, $col).Value2
            if ($val -ne $null -and $val.ToString().Trim() -ne "") {
                Write-Host "Col $col ($([char](64 + $col))): $val"
            }
        }
    } catch {
        Write-Host "Error accessing $sheetName : $_"
    }
}

$gijonName = "RESIDENCIA GIJ" + [char]211 + "N"
DumpHeaders($gijonName)
DumpHeaders("RESIDENCIA SOTO")
DumpHeaders("RESIDENCIA OVIEDO")

$wb.Close($false)
$xl.Quit()
