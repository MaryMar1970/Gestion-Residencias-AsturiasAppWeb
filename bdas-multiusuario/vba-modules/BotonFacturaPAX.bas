Attribute VB_Name = "BotonFacturaPAX"
'Attribute VB_Name = "BotonFacturaPAX"
Option Explicit
'=================================================================================
' Modulo: BotonFacturaPAX
' Proposito: Botones para mostrar hojas de factura PAX por ubicacion
' Version: 1.2 - Parametro control opcional para maxima compatibilidad
' Fecha: 2025-01-14
'=================================================================================

' Constantes para nombres de hojas
Private Const HOJA_FACTURA_GIJON As String = "FACTURA GIJ�N"
Private Const HOJA_FACTURA_SOTO As String = "FACTURA SOTO"
Private Const HOJA_FACTURA_OVIEDO As String = "FACTURA OVIEDO"

' Procedimientos publicos para el Ribbon - parametro OPCIONAL
Public Sub MostrarFacturaPaxGijon(Optional control As IRibbonControl)
    Call MostrarHojaFacturaPax(HOJA_FACTURA_GIJON, "GIJ�N")
End Sub

Public Sub MostrarFacturaPaxSoto(Optional control As IRibbonControl)
    Call MostrarHojaFacturaPax(HOJA_FACTURA_SOTO, "SOTO")
End Sub

Public Sub MostrarFacturaPaxOviedo(Optional control As IRibbonControl)
    Call MostrarHojaFacturaPax(HOJA_FACTURA_OVIEDO, "OVIEDO")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub MostrarHojaFacturaPax(nombreHoja As String, ubicacion As String)
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
        Err.Raise vbObjectError + 6001, "MostrarHojaFacturaPax", _
                  "La hoja '" & nombreHoja & "' no existe en el libro"
    End If
    
    ' Mostrar y activar la hoja
    With ws
        .visible = xlSheetVisible
        .Activate
    End With
    
    Application.StatusBar = "Mostrando factura PAX de " & ubicacion
    
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
    mensajeError = "Error al mostrar factura PAX de " & ubicacion & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Hoja: " & nombreHoja
    
    Application.StatusBar = "ERROR: No se pudo mostrar la factura PAX"
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error al Mostrar Factura PAX"
    
    Resume CleanExit
End Sub

