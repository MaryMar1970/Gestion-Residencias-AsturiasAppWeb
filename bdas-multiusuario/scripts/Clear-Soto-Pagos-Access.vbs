Dim cn
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\ResidenciaBD\Residencia_BE.accdb;"
cn.Execute "UPDATE Ordenes SET EstadoPago = '' WHERE Residencia = 'SOTO';"
cn.Close
WScript.Echo "Campo EstadoPago limpiado para SOTO en Access."
