Attribute VB_Name = "BotonGestionPasswords"
'=================================================================================
' CALLBACK: Gesti�n de Contrase�as desde Ribbon
' Implementado: 2025-11-21 23:40 UTC
' Usuario: Bustiello2
'=================================================================================

Sub RibbonGestionPasswords(control As IRibbonControl)
    '=================================================================
    ' Callback para el bot�n "Gesti�n Contrase�as" del Ribbon
    ' Abre el formulario visual de gesti�n de contrase�as
    '=================================================================
    
    On Error GoTo ErrorHandler
    
    ' Abrir formulario de gesti�n
    Call ModuloPruebasPasswords.AbrirGestionPasswords
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al abrir la gesti�n de contrase�as:" & vbCrLf & vbCrLf & _
           "N�mero: " & Err.Number & vbCrLf & _
           "Descripci�n: " & Err.Description, _
           vbCritical, "Error de Gesti�n de Contrase�as"
End Sub

