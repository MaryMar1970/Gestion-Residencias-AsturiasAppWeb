Attribute VB_Name = "BotonBloqueoHabitaciones"
'Attribute VB_Name = "BotonBloqueoHabitaciones"
Option Explicit
'=================================================================================
' Modulo: BotonBloqueoHabitaciones
' Proposito: Botones para mostrar hojas de bloqueos por ubicacion
' Version: 1.0 - Consolidado y Robusto
' Fecha: 2025-01-14
'=================================================================================

' Constantes para nombres de hojas
Private Const HOJA_BLOQUEOS_GIJON As String = "Habitaciones Bloqueadas GIJ�N"
Private Const HOJA_BLOQUEOS_SOTO As String = "Apartamentos Bloqueados SOTO"
Private Const HOJA_BLOQUEOS_OVIEDO As String = "Habitaciones Bloqueadas OVIEDO"

' Procedimientos publicos para el Ribbon
Public Sub MostrarBloqueosGijon(control As IRibbonControl)
    Call MostrarHojaBloqueos(HOJA_BLOQUEOS_GIJON, "GIJ�N")
End Sub

Public Sub MostrarBloqueosSoto(control As IRibbonControl)
    Call MostrarHojaBloqueos(HOJA_BLOQUEOS_SOTO, "SOTO")
End Sub

Public Sub MostrarBloqueosOviedo(control As IRibbonControl)
    Call MostrarHojaBloqueos(HOJA_BLOQUEOS_OVIEDO, "OVIEDO")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub MostrarHojaBloqueos(nombreHoja As String, ubicacion As String)
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
        Err.Raise vbObjectError + 2001, "MostrarHojaBloqueos", _
                  "La hoja '" & nombreHoja & "' no existe en el libro"
    End If
    
    ' Mostrar y activar la hoja
    With hoja
        .visible = xlSheetVisible
        .Activate
    End With
    
    Application.StatusBar = "Mostrando bloqueos de " & ubicacion
    
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
    mensajeError = "Error al mostrar hoja de bloqueos de " & ubicacion & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Hoja: " & nombreHoja
    
    Application.StatusBar = "ERROR: No se pudo mostrar la hoja de bloqueos"
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error al Mostrar Bloqueos"
    
    Resume CleanExit
End Sub

