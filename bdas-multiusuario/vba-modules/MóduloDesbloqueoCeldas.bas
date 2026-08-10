Attribute VB_Name = "M�duloDesbloqueoCeldas"
Sub DesbloquearTodasLasHojas()
    Dim ws As Worksheet
    Dim clave As String

    ' Solicitar contrase�a si es necesaria
    clave = InputBox("Introduce la contrase�a para desbloquear las hojas (deja en blanco si no tienen contrase�a):", "Desbloquear hojas")

    On Error Resume Next ' Evita error si una hoja no tiene protecci�n o clave incorrecta

    For Each ws In ThisWorkbook.Worksheets
        ws.Unprotect password:=clave
    Next ws

    On Error GoTo 0 ' Restaurar manejo normal de errores

    MsgBox "Proceso completado. Todas las hojas han sido desbloqueadas (si la contrase�a era correcta).", vbInformation
End Sub

