Attribute VB_Name = "GestionCacheCalendarios"
'Attribute VB_Name = "GestionCacheCalendarios"
Option Explicit

'=================================================================================
' Modulo: GestionCacheCalendarios
' Version: 2.0 - Cache por Fecha de Corte
' Fecha: 2025-12-04
' Proposito: Gestion centralizada de cache para los 3 calendarios
'            Permite mantener el historico "grabado" y solo actualizar futuras
'=================================================================================

' Variables privadas para control de cache - GIJON
Private cacheGijonActivo As Boolean
Private fechaCorteGijon As Date
Private ultimaActualizacionGijon As Double

' Variables privadas para control de cache - SOTO
Private cacheSotoActivo As Boolean
Private fechaCorteSoto As Date
Private ultimaActualizacionSoto As Double

' Variables privadas para control de cache - OVIEDO
Private cacheOviedoActivo As Boolean
Private fechaCorteOviedo As Date
Private ultimaActualizacionOviedo As Double

'=================================================================================
' FUNCIONES PUBLICAS - GIJON
'=================================================================================

Public Sub MarcarCacheGijonCargado()
    cacheGijonActivo = True
    fechaCorteGijon = Date
    ultimaActualizacionGijon = Timer
End Sub

Public Function CacheGijonCargado() As Boolean
    CacheGijonCargado = cacheGijonActivo
End Function

Public Function ObtenerFechaCorteGijon() As Date
    If cacheGijonActivo Then
        ObtenerFechaCorteGijon = fechaCorteGijon
    Else
        ObtenerFechaCorteGijon = 0
    End If
End Function

Public Sub LimpiarCacheGijon()
    cacheGijonActivo = False
    fechaCorteGijon = 0
    ultimaActualizacionGijon = 0
End Sub

Public Function TiempoDesdeActualizacionGijon() As Double
    If cacheGijonActivo Then
        TiempoDesdeActualizacionGijon = Timer - ultimaActualizacionGijon
    Else
        TiempoDesdeActualizacionGijon = 0
    End If
End Function

'=================================================================================
' FUNCIONES PUBLICAS - SOTO
'=================================================================================

Public Sub MarcarCacheSotoCargado()
    cacheSotoActivo = True
    fechaCorteSoto = Date
    ultimaActualizacionSoto = Timer
End Sub

Public Function CacheSotoCargado() As Boolean
    CacheSotoCargado = cacheSotoActivo
End Function

Public Function ObtenerFechaCorteSoto() As Date
    If cacheSotoActivo Then
        ObtenerFechaCorteSoto = fechaCorteSoto
    Else
        ObtenerFechaCorteSoto = 0
    End If
End Function

Public Sub LimpiarCacheSoto()
    cacheSotoActivo = False
    fechaCorteSoto = 0
    ultimaActualizacionSoto = 0
End Sub

Public Function TiempoDesdeActualizacionSoto() As Double
    If cacheSotoActivo Then
        TiempoDesdeActualizacionSoto = Timer - ultimaActualizacionSoto
    Else
        TiempoDesdeActualizacionSoto = 0
    End If
End Function

'=================================================================================
' FUNCIONES PUBLICAS - OVIEDO
'=================================================================================

Public Sub MarcarCacheOviedoCargado()
    cacheOviedoActivo = True
    fechaCorteOviedo = Date
    ultimaActualizacionOviedo = Timer
End Sub

Public Function CacheOviedoCargado() As Boolean
    CacheOviedoCargado = cacheOviedoActivo
End Function

Public Function ObtenerFechaCorteOviedo() As Date
    If cacheOviedoActivo Then
        ObtenerFechaCorteOviedo = fechaCorteOviedo
    Else
        ObtenerFechaCorteOviedo = 0
    End If
End Function

Public Sub LimpiarCacheOviedo()
    cacheOviedoActivo = False
    fechaCorteOviedo = 0
    ultimaActualizacionOviedo = 0
End Sub

Public Function TiempoDesdeActualizacionOviedo() As Double
    If cacheOviedoActivo Then
        TiempoDesdeActualizacionOviedo = Timer - ultimaActualizacionOviedo
    Else
        TiempoDesdeActualizacionOviedo = 0
    End If
End Function

'=================================================================================
' FUNCIONES GLOBALES
'=================================================================================

Public Sub LimpiarTodasLasCaches()
    LimpiarCacheGijon
    LimpiarCacheSoto
    LimpiarCacheOviedo
    Application.StatusBar = "Todas las caches limpiadas"
End Sub

Public Function ObtenerEstadoCaches() As String
    Dim estado As String
    estado = "ESTADO DE CACHES:" & vbCrLf & vbCrLf
    
    If CacheGijonCargado() Then
        estado = estado & "Gijon: Cargado (corte: " & Format(fechaCorteGijon, "dd/mm/yyyy") & _
                 " - " & Format(TiempoDesdeActualizacionGijon() / 60, "0.0") & " min)" & vbCrLf
    Else
        estado = estado & "Gijon: No cargado" & vbCrLf
    End If
    
    If CacheSotoCargado() Then
        estado = estado & "Soto: Cargado (corte: " & Format(fechaCorteSoto, "dd/mm/yyyy") & _
                 " - " & Format(TiempoDesdeActualizacionSoto() / 60, "0.0") & " min)" & vbCrLf
    Else
        estado = estado & "Soto: No cargado" & vbCrLf
    End If
    
    If CacheOviedoCargado() Then
        estado = estado & "Oviedo: Cargado (corte: " & Format(fechaCorteOviedo, "dd/mm/yyyy") & _
                 " - " & Format(TiempoDesdeActualizacionOviedo() / 60, "0.0") & " min)" & vbCrLf
    Else
        estado = estado & "Oviedo: No cargado" & vbCrLf
    End If
    
    ObtenerEstadoCaches = estado
End Function

