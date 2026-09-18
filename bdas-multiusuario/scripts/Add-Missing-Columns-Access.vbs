Dim cn, sqlList, i, sql
Set cn = CreateObject("ADODB.Connection")
cn.Open "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\ResidenciaBD\Residencia_BE.accdb;"

sqlList = Array( _
    "ALTER TABLE Ordenes ADD COLUMN Finalidad VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Empleo VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Situacion VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Evaluacion VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Comision VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Turno VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Rango VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN DiasUso INTEGER;", _
    "ALTER TABLE Ordenes ADD COLUMN PAX INTEGER;", _
    "ALTER TABLE Ordenes ADD COLUMN CamaSuple VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN DtoFamNum VARCHAR(50);", _
    "ALTER TABLE Ordenes ADD COLUMN Importe DOUBLE;", _
    "ALTER TABLE Ordenes ADD COLUMN FechaGrabacion VARCHAR(255);" _
)

For i = LBound(sqlList) To UBound(sqlList)
    sql = sqlList(i)
    On Error Resume Next
    cn.Execute sql
    If Err.Number = 0 Then
        WScript.Echo "OK: " & sql
    Else
        WScript.Echo "Nota (" & Err.Description & ") en: " & sql
    End If
    On Error GoTo 0
Next

cn.Close
WScript.Echo "MODIFICACION DE SCHEMA EN ACCESS COMPLETADA."
