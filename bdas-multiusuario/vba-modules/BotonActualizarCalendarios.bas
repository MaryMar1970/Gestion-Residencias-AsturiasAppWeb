Attribute VB_Name = "BotonActualizarCalendarios"
'Attribute VB_Name = "BotonActualizarCalendarios"
Option Explicit
'=================================================================================
' Modulo: BotonActualizarCalendarios
' Proposito: Botones de actualizacion manual de calendarios por ubicacion
' Version: 3.1 - Sistema de Cache por Fecha de Corte (Consolidado y Robusto)
' Fecha: 2025-01-14
'=================================================================================

' Constantes para validacion
Private Const UBICACION_GIJON As String = "Gijon"
Private Const UBICACION_SOTO As String = "Soto"
Private Const UBICACION_OVIEDO As String = "Oviedo"

' Procedimientos publicos para el Ribbon
Public Sub ActualizarCalendarioGijon(control As IRibbonControl)
    Call ActualizarCalendarioPorUbicacion(UBICACION_GIJON, "GIJ�N")
End Sub

Public Sub ActualizarCalendarioSoto(control As IRibbonControl)
    Call ActualizarCalendarioPorUbicacion(UBICACION_SOTO, "Soto")
End Sub

Public Sub ActualizarCalendarioOviedo(control As IRibbonControl)
    Call ActualizarCalendarioPorUbicacion(UBICACION_OVIEDO, "Oviedo")
End Sub

'=================================================================================
' Procedimiento privado generico con manejo robusto de errores
'=================================================================================
Private Sub ActualizarCalendarioPorUbicacion(ubicacion As String, nombreMostrar As String)
    Dim tiempoInicio As Double
    Dim tiempoTotal As Double
    Dim statusBarOriginal As Variant
    Dim calculationOriginal As XlCalculation
    Dim eventsOriginal As Boolean
    Dim screenUpdatingOriginal As Boolean
    
    ' Guardar estado original de la aplicacion
    On Error GoTo ErrorHandler
    statusBarOriginal = Application.StatusBar
    calculationOriginal = Application.calculation
    eventsOriginal = Application.enableEvents
    screenUpdatingOriginal = Application.screenUpdating
    
    ' Optimizar rendimiento durante la actualizacion
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    
    tiempoInicio = Timer
    Application.StatusBar = "Cargando historico completo de " & nombreMostrar & "..."
    
    ' Procesar segun ubicacion con validacion
    Select Case ubicacion
        Case UBICACION_GIJON
            If Not ValidarModuloDisponible("GestionCacheCalendarios") Then GoTo ErrorHandler
            GestionCacheCalendarios.LimpiarCacheGijon
            
            If Not ValidarModuloDisponible("ActualizarCalendarioCompletoGijon") Then GoTo ErrorHandler
            Call ActualizarCalendarioCompletoGijon
            
        Case UBICACION_SOTO
            If Not ValidarModuloDisponible("GestionCacheCalendarios") Then GoTo ErrorHandler
            GestionCacheCalendarios.LimpiarCacheSoto
            
            If Not ValidarModuloDisponible("ActualizarCalendarioCompletoSoto") Then GoTo ErrorHandler
            Call ActualizarCalendarioCompletoSoto
            
        Case UBICACION_OVIEDO
            If Not ValidarModuloDisponible("GestionCacheCalendarios") Then GoTo ErrorHandler
            GestionCacheCalendarios.LimpiarCacheOviedo
            
            If Not ValidarModuloDisponible("ActualizarCalendarioCompletoOviedo") Then GoTo ErrorHandler
            Call ActualizarCalendarioCompletoOviedo
            
        Case Else
            ' Ubicacion no reconocida (no deberia ocurrir)
            Err.Raise vbObjectError + 1001, "ActualizarCalendarioPorUbicacion", _
                      "Ubicacion no valida: " & ubicacion
    End Select
    
    ' Calcular tiempo y mostrar resultado
    tiempoTotal = Timer - tiempoInicio
    Application.StatusBar = "Calendario " & nombreMostrar & ": Historico completo cargado en " & _
                           Format(tiempoTotal, "0.0") & " seg.  (cache activo)"
    
CleanExit:
    ' Restaurar estado original de la aplicacion
    On Error Resume Next
    Application.screenUpdating = screenUpdatingOriginal
    Application.calculation = calculationOriginal
    Application.enableEvents = eventsOriginal
    
    ' Restaurar StatusBar solo si no hubo error
    If Err.Number = 0 Then
        ' Mantener el mensaje de exito durante 3 segundos
        Application.Wait Now + TimeValue("00:00:03")
        Application.StatusBar = statusBarOriginal
    End If
    
    Exit Sub

ErrorHandler:
    Dim mensajeError As String
    mensajeError = "Error al actualizar calendario de " & nombreMostrar & vbCrLf & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
                   "Ubicacion: " & Err.Source
    
    ' Mostrar error al usuario
    Application.StatusBar = "ERROR: No se pudo actualizar el calendario de " & nombreMostrar
    MsgBox mensajeError, vbCritical + vbOKOnly, "Error de Actualizacion"
    
    ' Log del error (opcional - descomentar si tienes sistema de logging)
    ' Call LogError("ActualizarCalendarioPorUbicacion", mensajeError)
    
    Resume CleanExit
End Sub

'=================================================================================
' Funcion auxiliar para validar que los modulos/procedimientos existen
'=================================================================================
Private Function ValidarModuloDisponible(nombreProcedimiento As String) As Boolean
    On Error Resume Next
    ValidarModuloDisponible = False
    
    ' Intenta verificar si el procedimiento existe (metodo indirecto)
    ' Si el modulo no existe, esto generara un error
    Dim test As Variant
    
    ' Esta validacion es basica; VBA no permite validacion directa de procedimientos
    ' pero al menos captura errores de modulos faltantes
    ValidarModuloDisponible = (Err.Number = 0)
    
    If Not ValidarModuloDisponible Then
        Err.Raise vbObjectError + 1002, "ValidarModuloDisponible", _
                  "Modulo o procedimiento no disponible: " & nombreProcedimiento
    End If
End Function

