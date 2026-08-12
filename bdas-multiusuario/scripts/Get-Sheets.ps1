$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $xl.Workbooks.Open('H:\ResidenciaBD\BDAS_Multiusuario.xlsm', 0, $true)
foreach ($s in $wb.Sheets) {
    Write-Host "Sheet Name: [$($s.Name)]"
}
$wb.Close($false)
$xl.Quit()
