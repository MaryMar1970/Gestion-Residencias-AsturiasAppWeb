$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.EnableEvents = $false
$wb = $xl.Workbooks.Open('H:\ResidenciaBD\BDAS_Multiusuario.xlsm', 0, $true)
Write-Host "ProtectStructure: $($wb.ProtectStructure)"
Write-Host "ProtectWindows  : $($wb.ProtectWindows)"
$wb.Close($false)
$xl.Quit()
