Attribute VB_Name = "ModuloRibbonRGPD"
' Attribute VB_Name = "ModuloRibbonRGPD"
'=================================================================================
' M�DULO: ModuloRibbonRGPD
' Prop�sito: Callbacks para botones del Ribbon personalizado RGPD
' Fecha: 2025-11-20 00:05
' Usuario: Bustiello2
'=================================================================================

Option Explicit

' Variable global para almacenar referencia al Ribbon
Public ribbonRGPD As IRibbonUI


'=================================================================================
' CALLBACK: Bot�n "Panel de Control" en pesta�a RGPD
'=================================================================================
Public Sub RibbonAbrirPanelRGPD(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' Mostrar formulario de control RGPD
    frmPanelRGPD.Show
    
    Exit Sub
    
ErrorHandler:
    MsgBox "? Error al abrir Panel RGPD:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' CALLBACK: Bot�n "Archivar Ahora" en pesta�a RGPD
'=================================================================================
Public Sub RibbonArchivarAhora(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' PASO 1: Contar cu�ntos registros se van a archivar
    Dim totalRegistros As Long
    totalRegistros = ModuloArchivadoDatosBDAS.ContarRegistrosParaArchivar()
    
    If totalRegistros = 0 Then
        MsgBox "? No hay registros antiguos para archivar." & vbCrLf & vbCrLf & _
               "Todos los registros tienen menos de 90 d�as desde la fecha de salida.", _
               vbInformation, "Archivado RGPD - Sin Acci�n Necesaria"
        Exit Sub
    End If
    
    ' PASO 2: Mostrar resumen y pedir confirmaci�n
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("SE VAN A ARCHIVAR " & totalRegistros & " REGISTROS" & vbCrLf & vbCrLf & _
                       "DATOS QUE SE ELIMINAR�N de hojas operativas:" & vbCrLf & _
                       "  � DNI / NIF" & vbCrLf & _
                       "  � Nombre completo" & vbCrLf & _
                       "  � Tel�fono" & vbCrLf & _
                       "  � Direcci�n" & vbCrLf & _
                       "  � Fecha grabaci�n solicitud" & vbCrLf & vbCrLf & _
                       "DATOS QUE SE CONSERVAN (RGPD Art. 5.1.b):" & vbCrLf & _
                       "  � N� Orden, Fechas estancia, Importes, Estado pago" & vbCrLf & vbCrLf & _
                       "Copia completa guardada en hoja BDAS" & vbCrLf & vbCrLf & _
                       "�Continuar con el archivado?", _
                       vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Archivado RGPD")
    
    If respuesta = vbNo Then
        MsgBox "Operaci�n cancelada por el usuario.", vbInformation, "Archivado RGPD"
        Exit Sub
    End If
    
    ' PASO 3: Ejecutar archivado manual
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar archivado:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error en Archivado RGPD"
End Sub


'=================================================================================
' CALLBACK: Bot�n "Ayuda" en pesta�a RGPD
' VERSI�N 2 MENSAJES: Sin emoticonos + Configuraci�n en mensaje 1
'=================================================================================
'Public Sub RibbonAyudaRGPD(control As IRibbonControl)
'    Dim respuesta As VbMsgBoxResult
'    Dim usuarioActual As String
'    Dim diasRetencion As String
'    Dim msg1 As String
'    Dim msg2 As String
    
'    Application.screenUpdating = True
'    DoEvents
    
    'Obtener datos
'    usuarioActual = Trim(ThisWorkbook.usuarioActual)
'    If usuarioActual = "" Then usuarioActual = "NO_IDENTIFICADO"
'    diasRetencion = ObtenerDiasRetencion()
    
    ' -------------------------------------------------------------------
    ' MENSAJE 1: Informaci�n pr�ctica y operativa
    ' -------------------------------------------------------------------
'    msg1 = "       SISTEMA DE ARCHIVADO AUTOM�TICO RGPD" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "NORMATIVA APLICABLE:" & vbCrLf
'    msg1 = msg1 & "� Reglamento UE 2016/679 (RGPD)" & vbCrLf
'    msg1 = msg1 & "  - Arts. 5.1.c y 5.1.e - Limitaci�n de conservaci�n" & vbCrLf
'    msg1 = msg1 & "� Real Decreto 933/2021, de 26 de octubre" & vbCrLf
'    msg1 = msg1 & "  - Conservaci�n: 3 a�os desde finalizaci�n servicio" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "FUNCIONAMIENTO ARCHIVADO:" & vbCrLf
'    msg1 = msg1 & "1. Autom�tico: Al abrir Excel, archiva > " & diasRetencion & " d�as" & vbCrLf
'    msg1 = msg1 & "2. Manual: Use 'Archivar Ahora'" & vbCrLf
'    msg1 = msg1 & "3. Criterio (configurable): " & diasRetencion & " d�as desde fecha SALIDA" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "DESTINO DEL ARCHIVADO:" & vbCrLf
'    msg1 = msg1 & "Cada residencia archiva en su propia BDAS." & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "DATOS ELIMINADOS / ARCHIVADOS(hojas operativas):" & vbCrLf
'    msg1 = msg1 & "DNI/NIF, Nombre, Tel�fono, Direcci�n, Grabaci�n" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "DATOS CONSERVADOS:" & vbCrLf
'    msg1 = msg1 & "N� Orden, N� Factura, Fechas estancia, Importes, Estado pago" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & "CONFIGURACI�N:" & vbCrLf
'    msg1 = msg1 & "Use 'Panel de Control' para:" & vbCrLf
'    msg1 = msg1 & "� Cambiar d�as de retenci�n (actual: " & diasRetencion & " d�as)" & vbCrLf
'    msg1 = msg1 & "� Activar/desactivar archivado autom�tico" & vbCrLf
'    msg1 = msg1 & "� Ver estad�sticas de archivado" & vbCrLf & vbCrLf
    
'    msg1 = msg1 & String(50, "-") & vbCrLf
'    msg1 = msg1 & "Usuario: " & usuarioActual & " | " & Format(Now, "dd/mm/yyyy hh:mm") & vbCrLf & vbCrLf
'    msg1 = msg1 & "�Desea ver informaci�n legal detallada?"
    
'    respuesta = MsgBox(msg1, vbInformation + vbYesNo, "Ayuda RGPD - Informaci�n General")
    
    ' -------------------------------------------------------------------
    ' MENSAJE 2: Informaci�n legal detallada (solo si usuario acepta)
    ' -------------------------------------------------------------------
'    If respuesta = vbYes Then
'        msg2 = "       INFORMACI�N LEGAL DETALLADA" & vbCrLf & vbCrLf
        
'        msg2 = msg2 & "FUNDAMENTO LEGAL - DATOS CONSERVADOS:" & vbCrLf & vbCrLf
'        msg2 = msg2 & "Art�culo 6. 1.c RGPD:" & vbCrLf
'        msg2 = msg2 & "Permite conservar datos cuando el tratamiento es" & vbCrLf
'        msg2 = msg2 & "necesario para cumplir una obligaci�n legal aplicable" & vbCrLf
'        msg2 = msg2 & "al responsable (obligaciones fiscales y contables)." & vbCrLf & vbCrLf
        
'        msg2 = msg2 & "Por ello se conservan:" & vbCrLf
'        msg2 = msg2 & "� N� de Orden" & vbCrLf
'        msg2 = msg2 & "� N� de Factura" & vbCrLf
'        msg2 = msg2 & "� Fechas de estancia (entrada/salida)" & vbCrLf
'        msg2 = msg2 & "� Importes y conceptos econ�micos" & vbCrLf
'        msg2 = msg2 & "� Estado de pago" & vbCrLf & vbCrLf
        
 '       msg2 = msg2 & "COPIA EN HOJA BDAS:" & vbCrLf & vbCrLf
 '       msg2 = msg2 & "� Los datos (incluidos personales) se guardan" & vbCrLf
 '       msg2 = msg2 & "� Acceso protegido con contrase�a" & vbCrLf
 '       msg2 = msg2 & "� Conservaci�n: m�mino 3 a�os desde fin del servicio" & vbCrLf & vbCrLf
        
 '       msg2 = msg2 & "Real Decreto 933/2021:" & vbCrLf
 '       msg2 = msg2 & "Este RD obliga a los centros residenciales a conservar" & vbCrLf
 '       msg2 = msg2 & "el registro inform�tico de usuarios durante al menos 3 a�os," & vbCrLf
 '       msg2 = msg2 & "justificando la conservaci�n de los datos para posibles solicitudes por parte de las autoridades competentes" & vbCrLf
 '       msg2 = msg2 & "si son requeridos." & vbCrLf & vbCrLf
        
 '       msg2 = msg2 & String(50, "-") & vbCrLf
 '       msg2 = msg2 & "Usuario: " & usuarioActual & " | " & Format(Now, "dd/mm/yyyy hh:mm")
        
 '       MsgBox msg2, vbInformation, "Ayuda RGPD - Informaci�n Legal"
 '   End If
    
 '   DoEvents
 '   RegistrarAccesoAyuda
'End Sub


'=================================================================================
' FUNCI�N AUXILIAR: Obtener d�as de retenci�n desde configuraci�n
'=================================================================================
Private Function ObtenerDiasRetencion() As String
    On Error Resume Next
    Dim ws As Worksheet
    Dim dias As Variant
    
    ' Intentar leer desde hoja de configuraci�n
    Set ws = ThisWorkbook.Sheets("CONFIG")
    
' ---- BORRAR, CODIGO LOCALIZAR ERRORES --------- ---- BORRAR, CODIGO LOCALIZAR ERRORES ---------
    
' ---- BORRAR, CODIGO LOCALIZAR ERRORES --------- ---- BORRAR, CODIGO LOCALIZAR ERRORES ---------
    
    If Not ws Is Nothing Then
        dias = ws.Range("B2").Value
        If IsNumeric(dias) And dias > 0 Then
            ObtenerDiasRetencion = CStr(dias)
            Exit Function
        End If
    End If
    
    ' Valor por defecto
    ObtenerDiasRetencion = "90"
    On Error GoTo 0
End Function

'=================================================================================
' FUNCI�N AUXILIAR: Registrar acceso a ayuda
'=================================================================================
Private Sub RegistrarAccesoAyuda()
    On Error GoTo ErrorHandler
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim usuarioActual As String
    Dim screenStatus As Boolean
    
    usuarioActual = Trim(ThisWorkbook.usuarioActual)
    If usuarioActual = "" Then Exit Sub
    
    screenStatus = Application.screenUpdating
    Application.screenUpdating = False
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            
            With wsLOG
                .Cells(ultimaFila, 1).Value = usuarioActual
                .Cells(ultimaFila, 2).Value = Now
                .Cells(ultimaFila, 3).Value = ""
                .Cells(ultimaFila, 4).Value = "AYUDA RGPD"
                .Cells(ultimaFila, 5).Value = ""
                .Cells(ultimaFila, 6).Value = ""
            End With
            
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
        On Error GoTo ErrorHandler
    Next i
    
    Application.screenUpdating = screenStatus
    Exit Sub
    
ErrorHandler:
    Application.screenUpdating = screenStatus
    Debug.Print "Error en RegistrarAccesoAyuda: " & Err.Description
End Sub


