Attribute VB_Name = "BotonLOG"
'Attribute VB_Name = "BotonLOG"
Option Explicit
'=================================================================================
' Modulo: BotonLOG
' Proposito: Botones para visualizar hojas de LOG por ubicacion
' Version: 1.0 - Consolidado y Robusto
' Fecha: 2025-01-14
'=================================================================================

' Constantes para nombres de hojas
Private Const HOJA_LOG_GIJON As String = "LOG_GIJ�N"
Private Const HOJA_LOG_SOTO As String = "LOG_SOTO"
Private Const HOJA_LOG_OVIEDO As String = "LOG_OVIEDO"

' Procedimientos publicos para el Ribbon
Public Sub VisualizarLogGijon(control As IRibbonControl)
    Call VisualizarHojaLog(HOJA_LOG_GIJON, "GIJ�N")
End Sub

Public Sub VisualizarLogSoto(control As IRibbonControl)
    Call VisualizarHojaLog(HOJA_LOG_SOTO, "SOTO")
End Sub

Public Sub VisualizarLogOviedo(control As IRibbonControl)
    Call VisualizarHojaLog(HOJA_LOG_OVIEDO, "OVIEDO")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub VisualizarHojaLog(nombreHoja As String, ubicacion As String)
    Dim hoja As Worksheet
    Dim screenUpdatingOriginal As Boolean
    
    On Error GoTo ErrorHandler
    
    ' Guardar estado original
    screenUpdatingOriginal = Application.screenUpdating
    Application.screenUpdating = False
    
    ' Validar que la hoja existe
    Set hoja = Nothing
    On Error Resume Next
    Set hoja = ThisWorkbook.Sheets(nombreHoja)
    On Error GoTo ErrorHandler
    
    If hoja Is Nothing Then
        Err.Raise vbObjectError + 3001, "VisualizarHojaLog", _
                  "La hoja '" & nombreHoja & "' no existe en el libro"
    End If
    
    ' Mostrar y activar la hoja
    With hoja
        .visible = xlSheetVisible
        .Activate
    End With
    
    Application.StatusBar = "Mostrando LOG de " & ubicacion
    
CleanExit:
    ' Restaurar estado original
    On Error Resume Next
    Application.screenUpdating = screenUpdatingOriginal
    
    ' Limpiar StatusBar despues de 2 segundos
    If Err.Number = 0 Then
        Application.Wait Now + TimeValue("00:00:02")
        Application.StatusBar = False
    End If
    
    Exit Sub

ErrorHandler:
    Dim mensajeError As String
    mensajeError = "Error al visualizar LOG de " & ubicacion & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Hoja: " & nombreHoja
    
    Application.StatusBar = "ERROR: No se pudo visualizar el LOG"
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error al Visualizar LOG"
    
    Resume CleanExit
End Sub

