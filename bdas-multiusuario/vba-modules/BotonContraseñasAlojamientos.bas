Attribute VB_Name = "BotonContrase�asAlojamientos"
'Attribute VB_Name = "BotonContrase�asAlojamientos"
Option Explicit
'=================================================================================
' Modulo: BotonContrase�asAlojamientos
' Proposito: Botones para mostrar hojas de contrase�as por ubicacion
' Version: 1.0 - Consolidado y Robusto
' Fecha: 2025-01-14
'=================================================================================

' Constantes para nombres de hojas
Private Const HOJA_CONTRASENAS_GIJON As String = "CONTRASE�AS GIJ�N"
Private Const HOJA_CONTRASENAS_SOTO As String = "CONTRASE�AS SOTO"

' Procedimientos publicos para el Ribbon
Public Sub MostrarContrasenasGijon(control As IRibbonControl)
    Call MostrarHojaContrasenas(HOJA_CONTRASENAS_GIJON, "GIJ�N")
End Sub

Public Sub MostrarContrasenasSoto(control As IRibbonControl)
    Call MostrarHojaContrasenas(HOJA_CONTRASENAS_SOTO, "SOTO")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub MostrarHojaContrasenas(nombreHoja As String, ubicacion As String)
    Dim ws As Worksheet
    Dim screenUpdatingOriginal As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Guardar estado original
    screenUpdatingOriginal = Application.screenUpdating
    Application.screenUpdating = False
    
    ' Validar que la hoja existe
    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Err.Raise vbObjectError + 5001, "MostrarHojaContrasenas", _
                  "La hoja '" & nombreHoja & "' no existe en el libro"
    End If
    
    ' Mostrar y activar la hoja
    With ws
        .visible = xlSheetVisible
        .Activate
    End With
    
    Application.StatusBar = "Mostrando contrase�as de " & ubicacion
    
CleanExit:
    ' Restaurar estado original
    On Error Resume Next
    Application.screenUpdating = screenUpdatingOriginal
    
    ' Limpiar StatusBar despues de 2 segundos
    If Err.Number = 0 Then
        Application.Wait Now + TimeValue("00:00:02")
        Application.StatusBar = False
    End If
    
    ' Limpiar objetos
    Set ws = Nothing
    
    Exit Sub

ErrorHandler:
    Dim mensajeError As String
    mensajeError = "Error al mostrar contrase�as de " & ubicacion & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Hoja: " & nombreHoja
    
    Application.StatusBar = "ERROR: No se pudo mostrar la hoja de contrase�as"
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error al Mostrar Contrase�as"
    
    Resume CleanExit
End Sub

