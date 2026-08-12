$wsh = New-Object -ComObject WScript.Shell
$shortcut = $wsh.CreateShortcut("C:\Users\jaime_ntymfau\Desktop\BDAS Multiusuario.lnk")
Write-Host "Target Path: $($shortcut.TargetPath)"
Write-Host "Arguments  : $($shortcut.Arguments)"
Write-Host "Working Dir: $($shortcut.WorkingDirectory)"
