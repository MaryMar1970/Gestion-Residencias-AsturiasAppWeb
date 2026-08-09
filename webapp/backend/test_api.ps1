$body = @{email='admin@residencia.local'; password='Admin123!'} | ConvertTo-Json
$resp = Invoke-RestMethod -Uri 'http://localhost:5260/auth/login' -Method POST -Body $body -ContentType 'application/json'
$token = $resp.token
Write-Output "TOKEN obtained: $($token.Substring(0,20))..."

$headers = @{Authorization="Bearer $token"}
$habResp = Invoke-RestMethod -Uri 'http://localhost:5260/api/habitaciones' -Headers $headers
Write-Output "HABITACIONES COUNT: $($habResp.Count)"
$habResp | ForEach-Object { Write-Output "  HAB: $($_.numero) - $($_.residenciaNombre) - Activa: $($_.activa)" }

$calResp = Invoke-RestMethod -Uri 'http://localhost:5260/api/calendario?fechaInicio=2026-08-01&fechaFin=2026-08-15' -Headers $headers
Write-Output "CALENDARIO HABITACIONES COUNT: $($calResp.habitaciones.Count)"
$calResp.habitaciones | ForEach-Object { Write-Output "  CAL HAB: $($_.habitacionId) - $($_.numero) - $($_.residenciaNombre) - Activa: $($_.activa)" }
