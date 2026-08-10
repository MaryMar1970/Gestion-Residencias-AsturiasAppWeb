Attribute VB_Name = "ModuloGestionCacheCalendarios"
' Attribute VB_Name = "ModuloGestionCacheCalendarios"
'=================================================================================
' M�dulo: ModuloGestionCacheCalendarios
' Prop�sito: Gestionar el estado de cach� de los calendarios durante la sesi�n
' Fecha: 2025-11-17
' Autor: Bustiello
'=================================================================================

Option Explicit

' Variables de estado de cach� por residencia
Private historicoGijonCargado As Boolean
Private historicoOviedoCargado As Boolean
Private historicoSotoCargado As Boolean

' Marca de tiempo de �ltima actualizaci�n
Private ultimaActualizacionGijon As Date
Private ultimaActualizacionOviedo As Date
Private ultimaActualizacionSoto As Date

'=================================================================================
' FUNCIONES P�BLICAS - VERIFICAR CACH�
'=================================================================================

Public Function CacheGijonCargado() As Boolean
    CacheGijonCargado = historicoGijonCargado
End Function

Public Function CacheOviedoCargado() As Boolean
    CacheOviedoCargado = historicoOviedoCargado
End Function

Public Function CacheSotoCargado() As Boolean
    CacheSotoCargado = historicoSotoCargado
End Function

'=================================================================================
' FUNCIONES P�BLICAS - MARCAR CACH� COMO CARGADO
'=================================================================================

Public Sub MarcarCacheGijonCargado()
    historicoGijonCargado = True
    ultimaActualizacionGijon = Now
End Sub

Public Sub MarcarCacheOviedoCargado()
    historicoOviedoCargado = True
    ultimaActualizacionOviedo = Now
End Sub

Public Sub MarcarCacheSotoCargado()
    historicoSotoCargado = True
    ultimaActualizacionSoto = Now
End Sub

'=================================================================================
' FUNCIONES P�BLICAS - LIMPIAR CACH�
'=================================================================================

Public Sub LimpiarCacheGijon()
    historicoGijonCargado = False
    ultimaActualizacionGijon = 0
End Sub

Public Sub LimpiarCacheOviedo()
    historicoOviedoCargado = False
    ultimaActualizacionOviedo = 0
End Sub

Public Sub LimpiarCacheSoto()
    historicoSotoCargado = False
    ultimaActualizacionSoto = 0
End Sub

Public Sub LimpiarTodosCaches()
    LimpiarCacheGijon
    LimpiarCacheOviedo
    LimpiarCacheSoto
End Sub

'=================================================================================
' FUNCIONES P�BLICAS - OBTENER TIEMPO DESDE �LTIMA ACTUALIZACI�N
'=================================================================================

Public Function TiempoDesdeActualizacionGijon() As Double
    If ultimaActualizacionGijon = 0 Then
        TiempoDesdeActualizacionGijon = 999999
    Else
        TiempoDesdeActualizacionGijon = DateDiff("s", ultimaActualizacionGijon, Now)
    End If
End Function

Public Function TiempoDesdeActualizacionOviedo() As Double
    If ultimaActualizacionOviedo = 0 Then
        TiempoDesdeActualizacionOviedo = 999999
    Else
        TiempoDesdeActualizacionOviedo = DateDiff("s", ultimaActualizacionOviedo, Now)
    End If
End Function

Public Function TiempoDesdeActualizacionSoto() As Double
    If ultimaActualizacionSoto = 0 Then
        TiempoDesdeActualizacionSoto = 999999
    Else
        TiempoDesdeActualizacionSoto = DateDiff("s", ultimaActualizacionSoto, Now)
    End If
End Function
