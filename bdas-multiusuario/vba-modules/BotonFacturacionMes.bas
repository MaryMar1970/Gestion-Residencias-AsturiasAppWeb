Attribute VB_Name = "BotonFacturacionMes"
'Attribute VB_Name = "BotonFacturacionMes"
Option Explicit
'=================================================================================
' Modulo: BotonFacturacionMes
' Proposito: Botones para mostrar hojas de facturacion mensual por ubicacion
' Version: 1.0 - Consolidado y Robusto
' Fecha: 2025-01-14
'=================================================================================

' Constantes para nombres de hojas
Private Const HOJA_FACTURACION_GIJON As String = "FACTURACI�N GIJ�N"
Private Const HOJA_FACTURACION_SOTO As String = "FACTURACI�N SOTO"
Private Const HOJA_FACTURACION_OVIEDO As String = "FACTURACI�N OVIEDO"

' Procedimientos publicos para el Ribbon
Public Sub MostrarFacturacionGijon(control As IRibbonControl)
    Call MostrarHojaFacturacion(HOJA_FACTURACION_GIJON, "GIJ�N")
End Sub

Public Sub MostrarFacturacionSoto(control As IRibbonControl)
    Call MostrarHojaFacturacion(HOJA_FACTURACION_SOTO, "SOTO")
End Sub

Public Sub MostrarFacturacionOviedo(control As IRibbonControl)
    Call MostrarHojaFacturacion(HOJA_FACTURACION_OVIEDO, "OVIEDO")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub MostrarHojaFacturacion(nombreHoja As String, ubicacion As String)
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
        Err.Raise vbObjectError + 4001, "MostrarHojaFacturacion", _
                  "La hoja '" & nombreHoja & "' no existe en el libro"
    End If
    
    ' Mostrar y activar la hoja
    With ws
        .visible = xlSheetVisible
        .Activate
    End With
    
    Application.StatusBar = "Mostrando facturaci�n de " & ubicacion
    
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
    mensajeError = "Error al mostrar facturaci�n de " & ubicacion & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Hoja: " & nombreHoja
    
    Application.StatusBar = "ERROR: No se pudo mostrar la hoja de facturaci�n"
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error al Mostrar Facturaci�n"
    
    Resume CleanExit
End Sub

