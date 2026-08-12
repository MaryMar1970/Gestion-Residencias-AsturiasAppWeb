$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false

function CheckWb($path) {
    Write-Host "=== CHECKING: $path ==="
    $wb = $xl.Workbooks.Open($path, 0, $true)
    $comp = $wb.VBProject.VBComponents.Item("modDatabase")
    $cm = $comp.CodeModule
    $linesCount = $cm.CountOfLines
    Write-Host "modDatabase total lines: ${linesCount}"

    for ($i = 1; $i -le $linesCount; $i++) {
        $line = $cm.Lines($i, 1)
        if ($line -like "*ObtenerValorCampo*") {
            Write-Host "Line ${i}: ${line}"
        }
    }
    $wb.Close($false)
}

CheckWb('H:\ResidenciaApp\bdas-multiusuario\2026-08-07, BDAS_GIJ' + [char]211 + 'N-SOTO-OVIEDO_v16.8.4.xlsm')
CheckWb('H:\ResidenciaBD\BDAS_Multiusuario.xlsm')

$xl.Quit()
