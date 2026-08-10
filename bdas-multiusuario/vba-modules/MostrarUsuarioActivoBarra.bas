Attribute VB_Name = "MostrarUsuarioActivoBarra"
Sub MostrarUsuarioEnBarraEstado()
    Dim usuarioActual As String
    usuarioActual = Trim(ThisWorkbook.usuarioActual)
    
    If usuarioActual <> "" Then
        Application.StatusBar = ">>> Usuario activo: " & usuarioActual
    Else
        Application.StatusBar = "Usuario no identificado - Inicie sesi�n"
    End If
End Sub
