Attribute VB_Name = "CalendariosFechaInicio"
Option Explicit

'=================================================================================
' Modulo: CalendariosFechaInicio
' Proposito: Gesti�n centralizada de fechas de inicio de calendarios
' - Actualizaci�n autom�tica (Hoy - 100 d�as)
' - Actualizaci�n manual (fecha personalizada)
'=================================================================================

' ================================================================================
' ACTUALIZACI�N AUTOM�TICA - Llamada desde Workbook_Open
' ================================================================================
Public Sub AutoActualizarFechasCalendarios()
    ' Establece autom�ticamente: Hoy - 100 d�as en todos los calendarios
    ' Llamar desde Workbook_Open para actualizaci�n autom�tica al abrir el libro
    
    On Error GoTo ErrHandler
    
    Dim fechaInicio As Date
    fechaInicio = Date - 100  ' Hoy menos 100 d�as
    
    Application.StatusBar = "Actualizando fechas de calendarios (�ltimos 100 d�as)..."
    Application.screenUpdating = False
    Application.enableEvents = False
    
    ' Actualizar cada calendario
    Call ActualizarFechaCalendarioInterno("Calendario GIJ�N", fechaInicio)
    Call ActualizarFechaCalendarioInterno("Calendario SOTO", fechaInicio)
    Call ActualizarFechaCalendarioInterno("Calendario OVIEDO", fechaInicio)
    
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.StatusBar = False
    
    Exit Sub

ErrHandler:
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.StatusBar = False
    MsgBox "Error en actualizaci�n autom�tica: " & Err.Description, vbCritical, "Error"
End Sub

' ================================================================================
' ACTUALIZACI�N MANUAL - Bot�n Ribbon con confirmaciones
' ================================================================================
Public Sub EstablecerFechaInicio(control As IRibbonControl)
    On Error GoTo ErrHandler
    
    Dim ws As Worksheet
    Dim wsRes As Worksheet
    Dim nombreHojaRes As String
    Set ws = ActiveSheet
    
    ' Verificar hoja permitida
    Select Case ws.Name
        Case "Calendario GIJ�N"
            nombreHojaRes = "RESIDENCIA GIJ�N"
        Case "Calendario SOTO"
            nombreHojaRes = "RESIDENCIA SOTO"
        Case "Calendario OVIEDO"
            nombreHojaRes = "RESIDENCIA OVIEDO"
        Case Else
            MsgBox "Este bot�n solo funciona en hojas de calendario", vbExclamation, "Ubicaci�n incorrecta"
            Exit Sub
    End Select
    
    ' --- PRIMERA CONFIRMACI�N ---
    Dim confirm1 As VbMsgBoxResult
    confirm1 = MsgBox("�Desea establecer una nueva fecha de inicio del calendario?", _
                      vbQuestion + vbYesNo, "Confirmaci�n Inicial")
    
    If confirm1 <> vbYes Then Exit Sub
    
    ' --- SOLICITUD DE FECHA ---
    Dim fechaInput As String
    Dim fechaValida As Date
    Dim fechaOk As Boolean
    
    Do While Not fechaOk
        fechaInput = InputBox("Ingrese la nueva fecha de inicio (dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                              "O deje en blanco para usar: Hoy - 100 d�as", "Fecha de Inicio")
        
        ' Si usuario cancela completamente (presiona Cancelar)
        If StrPtr(fechaInput) = 0 Then Exit Sub
        
        ' Si deja en blanco (presiona OK sin escribir nada)
        If Trim(fechaInput) = "" Then
            ' Usar fecha autom�tica (Hoy - 100 d�as)
            fechaValida = Date - 100
            fechaOk = True
        ElseIf IsDate(fechaInput) Then
            fechaValida = CDate(fechaInput)
            fechaOk = True
        Else
            MsgBox "Formato de fecha inv�lido. Use dd/mm/aaaa", vbExclamation, "Error"
        End If
    Loop
    
    ' --- VERIFICAR RESERVAS ANTERIORES A LA NUEVA FECHA ---
    Dim reservasAnteriores As Long
    Dim fechaMinima As Date
    reservasAnteriores = 0
    fechaMinima = fechaValida
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets(nombreHojaRes)
    On Error GoTo ErrHandler
    
    If Not wsRes Is Nothing Then
        Dim ultimaFila As Long
        Dim fila As Long
        Dim fechaEntrada As Date
        Dim fechaSalida As Date
        
        ultimaFila = wsRes.Cells(wsRes.Rows.Count, "L").End(xlUp).Row
        
        For fila = 2 To ultimaFila
            If IsDate(wsRes.Cells(fila, "L").Value) And IsDate(wsRes.Cells(fila, "M").Value) Then
                fechaEntrada = CDate(wsRes.Cells(fila, "L").Value)
                fechaSalida = CDate(wsRes.Cells(fila, "M").Value)
                
                If fechaEntrada < fechaValida And fechaSalida >= fechaValida Then
                    reservasAnteriores = reservasAnteriores + 1
                    If fechaEntrada < fechaMinima Then
                        fechaMinima = fechaEntrada
                    End If
                End If
            End If
        Next fila
    End If
    
    ' --- AVISAR SI HAY RESERVAS QUE QUEDAR�AN CORTADAS ---
    If reservasAnteriores > 0 Then
        Dim msgReservas As String
        msgReservas = "ATENCI�N: Hay " & reservasAnteriores & " reserva(s) con fecha de entrada anterior al " & Format(fechaValida, "dd/mm/yyyy") & _
                      " pero que contin�an activas despu�s de esa fecha." & vbCrLf & vbCrLf & _
                      "Estas reservas se mostrar�n PARCIALMENTE en el calendario (solo la parte visible)." & vbCrLf & vbCrLf & _
                      "Fecha de entrada m�s antigua afectada: " & Format(fechaMinima, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                      "�Desea continuar de todos modos?"
        
        Dim confirmReservas As VbMsgBoxResult
        confirmReservas = MsgBox(msgReservas, vbExclamation + vbYesNo, "Reservas Afectadas")
        
        If confirmReservas <> vbYes Then Exit Sub
    End If
    
    ' --- SEGUNDA CONFIRMACI�N ---
    Dim confirm2 As VbMsgBoxResult
    confirm2 = MsgBox("Se proceder� a asignar nuevas fechas al calendario." & vbCrLf & vbCrLf & _
                      "Nueva fecha de inicio: " & Format(fechaValida, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                      "�Desea continuar?", _
                      vbExclamation + vbYesNo, "Confirmaci�n Final")
    
    If confirm2 <> vbYes Then Exit Sub
    
    ' --- APLICAR NUEVA FECHA ---
    Application.screenUpdating = False
    Application.enableEvents = False
    
    Call ActualizarFechaCalendarioInterno(ws.Name, fechaValida)
    
    ' --- ACTUALIZAR CALENDARIO COMPLETO ---
    Select Case ws.Name
        Case "Calendario GIJ�N"
            Application.enableEvents = True
            Application.screenUpdating = True
            Call ActualizarCalendarioFuturoGijon
        Case "Calendario SOTO"
            Application.enableEvents = True
            Application.screenUpdating = True
            Call ActualizarCalendarioFuturoSoto
        Case "Calendario OVIEDO"
            Application.enableEvents = True
            Application.screenUpdating = True
            Call ActualizarCalendarioFuturoOviedo
    End Select
    
    ' --- MENSAJE FINAL ---
    MsgBox "Fecha de inicio actualizada correctamente:" & vbCrLf & _
           Format(fechaValida, "dddd, dd mmmm yyyy") & vbCrLf & vbCrLf & _
           "El calendario ha sido regenerado y repintado.", _
           vbInformation, "Operaci�n Exitosa"
    
    Exit Sub

ErrHandler:
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "Error al establecer fecha: " & Err.Description, vbCritical, "Error"
End Sub

' ================================================================================
' FUNCI�N INTERNA - Actualiza fechas en un calendario espec�fico
' ================================================================================
Public Sub ActualizarFechaCalendarioInterno(nombreHoja As String, fechaInicio As Date)
    ' Funci�n auxiliar com�n para actualizaci�n autom�tica y manual
    ' No debe llamarse directamente desde fuera del m�dulo
    
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    
    If ws Is Nothing Then Exit Sub
    
    Dim estabaProtegida As Boolean
    estabaProtegida = ws.ProtectContents
    If estabaProtegida Then ws.Unprotect password:=""
 
    ' Regenerar fechas en fila 1
    Dim ultimaCol As Long
    Dim colActual As Long
    Dim fechaActual As Date
    
    ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    fechaActual = fechaInicio
    
    For colActual = 2 To ultimaCol Step 3
        ws.Cells(1, colActual).Value = fechaActual
        ws.Cells(1, colActual).NumberFormat = "dddd, dd-mmm"
        fechaActual = fechaActual + 1
    Next colActual
        If estabaProtegida Then
        ws.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    
    ' Limpiar cache seg�n calendario
    Select Case nombreHoja
        Case "Calendario GIJ�N"
            GestionCacheCalendarios.LimpiarCacheGijon
        Case "Calendario SOTO"
            GestionCacheCalendarios.LimpiarCacheSoto
        Case "Calendario OVIEDO"
            GestionCacheCalendarios.LimpiarCacheOviedo
    End Select
End Sub

