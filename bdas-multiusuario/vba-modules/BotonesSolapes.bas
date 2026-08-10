Attribute VB_Name = "BotonesSolapes"
'Attribute VB_Name = "BotonesSolapes"
Option Explicit

'=================================================================================
' M�dulo: BotonesSolapes
' Descripci�n: Botones de Ribbon para gesti�n de solapes aceptados
'              Cubre: Gij�n (activar + visualizar)
'              Ampliable a Oviedo y Soto cuando sea necesario
'=================================================================================

Public Sub AceptarSolapeGijon_GrupoPorHabitacion(control As IRibbonControl)
    Dim hab As String
    hab = Trim$(InputBox("Habitaci�n (ej: 1, 2, 3, Of.1, Est.2):", "Aceptar solape GRUPO - GIJ�N"))
    If hab = "" Then Exit Sub
    
    Dim sDesde As String
    sDesde = InputBox("Fecha desde (dd/mm/aaaa):", "Aceptar solape GRUPO - GIJ�N")
    If Trim$(sDesde) = "" Or Not IsDate(sDesde) Then
        MsgBox "Fecha desde no v�lida.", vbExclamation, "Dato incorrecto"
        Exit Sub
    End If
    
    Dim sHasta As String
    sHasta = InputBox("Fecha hasta (dd/mm/aaaa):", "Aceptar solape GRUPO - GIJ�N")
    If Trim$(sHasta) = "" Or Not IsDate(sHasta) Then
        MsgBox "Fecha hasta no v�lida.", vbExclamation, "Dato incorrecto"
        Exit Sub
    End If
    
    Dim motivo As String
    motivo = InputBox("Motivo (opcional):", "Aceptar solape GRUPO - GIJ�N")
    
    Dim clave As String
    clave = ClaveSolapeGrupoGijon(hab)
         Call GuardarSolapeAceptadoGijon(clave, CDate(sDesde), CDate(sHasta), motivo)
    
    MsgBox "Solape de GRUPO aceptado y guardado:" & vbCrLf & clave & vbCrLf & vbCrLf & _
           "El calendario se actualizar� autom�ticamente.", vbInformation, "OK - GIJ�N"
    
    ' --- Repintar calendario para mostrar solape aceptado en AZUL ---
    Call ModuloCalendarioGijon.ActualizarCalendarioCompletoGijon
End Sub

Public Sub VerHojaSolapesAceptadosGijon(control As IRibbonControl)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SOLAPES_ACEPTADOS_GIJON")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No existe todav�a la hoja de solapes aceptados para GIJ�N." & vbCrLf & _
               "Acepta al menos un solape primero.", vbInformation, "Sin datos - GIJ�N"
        Exit Sub
    End If
    
    ws.visible = xlSheetVisible
    ws.Columns("A:F").AutoFit
    ws.Activate
End Sub

'---------------------------------------------------------------------------------
' OVIEDO - ACTIVAR SOLAPE
' Llamado desde: ButtonActivarSolapeOviedo ? onAction="AceptarSolapeOviedo_GrupoPorHabitacion"
'---------------------------------------------------------------------------------
Public Sub AceptarSolapeOviedo_GrupoPorHabitacion(control As IRibbonControl)
    Dim hab As String
    hab = Trim$(InputBox("Habitaci�n (ej: 1, 2, 3 ... 15):", "Aceptar solape GRUPO - OVIEDO"))
    If hab = "" Then Exit Sub
    
    Dim sDesde As String
    sDesde = InputBox("Fecha desde (dd/mm/aaaa):", "Aceptar solape GRUPO - OVIEDO")
    If Trim$(sDesde) = "" Or Not IsDate(sDesde) Then
        MsgBox "Fecha desde no v�lida.", vbExclamation, "Dato incorrecto"
        Exit Sub
    End If
    
    Dim sHasta As String
    sHasta = InputBox("Fecha hasta (dd/mm/aaaa):", "Aceptar solape GRUPO - OVIEDO")
    If Trim$(sHasta) = "" Or Not IsDate(sHasta) Then
        MsgBox "Fecha hasta no v�lida.", vbExclamation, "Dato incorrecto"
        Exit Sub
    End If
    
    Dim motivo As String
    motivo = InputBox("Motivo (opcional):", "Aceptar solape GRUPO - OVIEDO")
    
    Dim clave As String
    clave = ClaveSolapeGrupoOviedo(hab)
        Call GuardarSolapeAceptadoOviedo(clave, CDate(sDesde), CDate(sHasta), motivo)
    
    MsgBox "Solape de GRUPO aceptado y guardado:" & vbCrLf & clave & vbCrLf & vbCrLf & _
           "El calendario se actualizar� autom�ticamente.", vbInformation, "OK - OVIEDO"
    
    ' --- Repintar calendario para mostrar solape aceptado en AZUL ---
    Call ModuloCalendarioOviedo.ActualizarCalendarioCompletoOviedo
End Sub

'---------------------------------------------------------------------------------
' OVIEDO - VISUALIZAR HOJA DE SOLAPES
' Llamado desde: ButtonVisualizarSolapeOviedo ? onAction="VerHojaSolapesAceptadosOviedo"
'---------------------------------------------------------------------------------
Public Sub VerHojaSolapesAceptadosOviedo(control As IRibbonControl)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("SOLAPES_ACEPTADOS_OVIEDO")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "No existe todav�a la hoja de solapes aceptados para OVIEDO." & vbCrLf & _
               "Acepta al menos un solape primero.", vbInformation, "Sin datos - OVIEDO"
        Exit Sub
    End If
    
    ws.visible = xlSheetVisible
    ws.Columns("A:F").AutoFit
    ws.Activate
End Sub
