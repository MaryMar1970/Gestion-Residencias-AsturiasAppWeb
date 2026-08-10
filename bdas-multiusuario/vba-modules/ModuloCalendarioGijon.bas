Attribute VB_Name = "ModuloCalendarioGijon"
'Attribute VB_Name = "ModuloCalendarioGijon" v14.11.3
Option Explicit
Private Const COLOR_SOLAPE_ACEPTADO As Long = 15773696 ' RGB(176, 224, 230) = PowderBlue (azul claro)

'=================================================================================
' Modulo: ModuloCalendarioGijon
' Version: 17.0 - Sistema de Cache por Fecha de Corte
' Fecha: 2025-12-04
'=================================================================================

Private tipoSolape As Object
Private solapesAceptados As Object


Private Sub DefinirTipoSolape()
    If tipoSolape Is Nothing Then
        Set tipoSolape = CreateObject("Scripting.Dictionary")
    Else
        tipoSolape.RemoveAll
    End If
End Sub

'=================================================================================
' INTERFACES PUBLICAS
'=================================================================================

Public Sub ActualizarCalendarioCompletoGijon()
    ' Llamado desde boton Ribbon - Actualiza TODO (historico + futuro)
    ' Marca el cache como cargado
    ActualizarCalendarioInterno False, False
    GestionCacheCalendarios.MarcarCacheGijonCargado
    Application.StatusBar = "Calendario GIJ�N: Historico completo cargado y en cache"
    Call ResaltarFestivos
End Sub

Public Sub ActualizarCalendarioFuturoGijon()
    ' Llamado al activar hoja - Solo actualiza futuras
    ' Si hay cache, respeta el historico
    ActualizarCalendarioInterno True, GestionCacheCalendarios.CacheGijonCargado()
End Sub

'=================================================================================
' PROCEDIMIENTO PRINCIPAL
'=================================================================================

Private Sub ActualizarCalendarioInterno(ByVal soloFuturas As Boolean, ByVal usarCache As Boolean)
    Dim wsRes As Worksheet, wsCal As Worksheet, wsBloq As Worksheet
    Dim dateMapping As Object
    Dim reservasProcessadas As Collection
    Dim bloqueosProcessados As Collection
    Dim mapaOcupacion As Object
    Dim huboSolape As Boolean
    Dim reservasOcultas As Long
    Dim fechaCorte As Date
    Dim colCorte As Long
    Dim estabaProtegida As Boolean
    
    On Error GoTo ErrorHandler
    
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA GIJ�N")
    Set wsCal = ThisWorkbook.Worksheets("Calendario GIJ�N")
    Set wsBloq = ThisWorkbook.Worksheets("Habitaciones Bloqueadas GIJ�N")
    On Error GoTo ErrorHandler
    
    If wsRes Is Nothing Or wsCal Is Nothing Or wsBloq Is Nothing Then
        MsgBox "ERROR: No se encontraron las hojas necesarias.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Desproteger la hoja del calendario temporalmente para realizar la limpieza y pintado sin errores
    estabaProtegida = wsCal.ProtectContents
    If estabaProtegida Then wsCal.Unprotect password:=""
    
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    Set dateMapping = CrearMapeoFechas(wsCal)
    If dateMapping Is Nothing Or dateMapping.Count = 0 Then
        MsgBox "ERROR: No se pudo crear mapeo de fechas.", vbExclamation, "Error"
        GoTo CleanExit
    End If
    
    ' Determinar fecha de corte y columna de corte
    If usarCache Then
        fechaCorte = GestionCacheCalendarios.ObtenerFechaCorteGijon()
        If fechaCorte = 0 Then fechaCorte = Date
        fechaCorte = fechaCorte - 3  ' 3 dias de margen
    Else
        fechaCorte = 0  ' Sin corte, actualizar todo
    End If
    
    ' Obtener columna de corte si aplica
    colCorte = 0
    If fechaCorte > 0 Then
        If dateMapping.Exists(CLng(fechaCorte)) Then
            colCorte = dateMapping(CLng(fechaCorte))
        Else
            ' Buscar la fecha mas cercana
            Dim fechaBuscar As Date
            For fechaBuscar = fechaCorte To fechaCorte + 10
                If dateMapping.Exists(CLng(fechaBuscar)) Then
                    colCorte = dateMapping(CLng(fechaBuscar))
                    Exit For
                End If
            Next fechaBuscar
        End If
    End If
    
    ' Limpiar calendario (completo o parcial segun cache)
    If usarCache And colCorte > 0 Then
        LimpiarCalendarioParcial wsCal, colCorte
    Else
        LimpiarCalendarioRapido wsCal
        LimpiarBloqueos wsCal
    End If
    
    Set reservasProcessadas = ProcesarReservasEnMemoriaV13(wsRes, dateMapping, soloFuturas, fechaCorte)
    Set bloqueosProcessados = ProcesarBloqueosEnMemoriaV13(wsBloq, dateMapping, soloFuturas, fechaCorte)
    
    ' Contar reservas asignadas a habitaciones ocultas (filas ocultas en calendario)
    Dim listaOcultas As String: listaOcultas = ""
    reservasOcultas = ContarReservasEnHabitacionesOcultas(wsRes, wsCal, listaOcultas)
    
    DefinirTipoSolape
    Set solapesAceptados = CargarSolapesAceptadosGijon()
    Set mapaOcupacion = CreateObject("Scripting.Dictionary")
    
    huboSolape = AplicarReservasYBloqueosConTramosV16(wsCal, reservasProcessadas, bloqueosProcessados, mapaOcupacion, tipoSolape)
    
    RepintarReservasPagadas wsRes, wsCal, reservasProcessadas
    
        If reservasOcultas > 0 Then
        Dim mensajeOcultas As String
        If reservasOcultas = 1 Then
            mensajeOcultas = "ATENCION: Hay 1 reserva actual/futura asignada a una habitacion oculta (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Esta reserva NO se muestra en el calendario." & vbCrLf & vbCrLf & _
                           "Para verla, debes mostrar la fila correspondiente en 'Calendario GIJ�N'."
        Else
            mensajeOcultas = "ATENCION: Hay " & reservasOcultas & " reservas actuales/futuras asignadas a habitaciones ocultas (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Estas reservas NO se muestran en el calendario." & vbCrLf & vbCrLf & _
                           "Para verlas, debes mostrar las filas correspondientes en 'Calendario GIJ�N'."
        End If
        MsgBox mensajeOcultas, vbExclamation, "Reservas en Habitaciones Ocultas"
    End If

CleanExit:
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    Exit Sub

ErrorHandler:
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "ERROR: " & Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' LIMPIEZA DE CALENDARIO
'=================================================================================

Private Sub LimpiarCalendarioRapido(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long
    Dim rango As Range
    
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
  
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub
    
    Set rango = wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol))
    
    ' Deshacer combinaciones
    On Error Resume Next
    rango.UnMerge
    On Error GoTo 0
    
    ' Limpiar contenido y formato
    rango.ClearContents
    rango.Interior.ColorIndex = xlNone
    rango.Font.Size = 11
    rango.Font.Bold = False
    Set rango = Nothing
End Sub

Private Sub LimpiarCalendarioParcial(wsCal As Worksheet, ByVal colDesde As Long)
    ' Limpia solo desde la columna indicada en adelante (para cache)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colDesde < 2 Then colDesde = 2
    
    Dim ultimaFila As Long, ultimaCol As Long
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Or ultimaCol < colDesde Then Exit Sub
    
    wsCal.Range(wsCal.Cells(2, colDesde), wsCal.Cells(ultimaFila, ultimaCol)).ClearContents
    wsCal.Range(wsCal.Cells(2, colDesde), wsCal.Cells(ultimaFila, ultimaCol)).Interior.ColorIndex = xlNone
End Sub

Private Sub LimpiarBloqueos(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long, i As Long, j As Long
    Dim arrValues As Variant
    
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub
    
    ' Leer toda la cuadr�cula en memoria de un solo golpe
    arrValues = wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol)).Value
    If Not IsArray(arrValues) Then Exit Sub
    
    For i = 2 To ultimaFila
        For j = 2 To ultimaCol
            Dim val As Variant
            val = arrValues(i - 1, j - 1)
            
            ' Solo consultamos el objeto Celda si contiene "BLOQUEADA" o si no est� vac�o
            If val = "BLOQUEADA" Or Not IsEmpty(val) Then
                If wsCal.Cells(i, j).Value = "BLOQUEADA" Or wsCal.Cells(i, j).Interior.color = RGB(255, 192, 0) Then
                    wsCal.Cells(i, j).Value = ""
                    wsCal.Cells(i, j).Interior.ColorIndex = xlNone
                End If
            End If
        Next j
    Next i
End Sub

'=================================================================================
' CREAR MAPEO DE FECHAS
'=================================================================================

Private Function CrearMapeoFechas(wsCal As Worksheet) As Object
    On Error Resume Next
    Set CrearMapeoFechas = Nothing
    If wsCal Is Nothing Then Exit Function
    
    Dim dateMapping As Object
    Set dateMapping = CreateObject("Scripting.Dictionary")
    Dim ultimaCol As Long, col As Long
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaCol < 2 Then Set CrearMapeoFechas = dateMapping: Exit Function
    
    For col = 2 To ultimaCol Step 3
        If IsDate(wsCal.Cells(1, col).Value) Then
            Dim fechaLng As Long
            fechaLng = CLng(wsCal.Cells(1, col).Value)
            If Not dateMapping.Exists(fechaLng) Then dateMapping.Add fechaLng, col
        End If
    Next col
    
    Set CrearMapeoFechas = dateMapping
End Function

'=================================================================================
' PROCESADO EN MEMORIA
'=================================================================================

Private Function ProcesarReservasEnMemoriaV13(wsRes As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    Dim reservasProcessadas As Collection
    Set reservasProcessadas = New Collection
    On Error Resume Next
    
    If wsRes Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    Dim arrRes As Variant
    arrRes = wsRes.Range("A2:S" & lastRow).Value
    If Not IsArray(arrRes) Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    Dim i As Long, room As String, arrRooms As Variant, r As Variant, reservaArray As Variant
    Dim hoy As Date: hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro
    If fechaCorte > 0 Then
        fechaFiltro = fechaCorte
    ElseIf soloFuturas Then
        fechaFiltro = hoy - 3
    Else
        fechaFiltro = 0
    End If
    
    For i = 1 To UBound(arrRes, 1)
        If IsDate(arrRes(i, 12)) And IsDate(arrRes(i, 13)) And Not IsEmpty(arrRes(i, 19)) Then
            ' Filtrar por fecha si aplica
            If fechaFiltro > 0 Then
                If CDate(arrRes(i, 13)) < fechaFiltro Then GoTo NextI
            End If
            
            room = Trim(CStr(arrRes(i, 19)))
            If InStr(room, ",") > 0 Then
                arrRooms = Split(room, ",")
                For Each r In arrRooms
                    If Trim(r) <> "" Then
                        reservaArray = CrearReservaArrayV12(arrRes, i, Trim(r), dateMapping)
                        If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
                    End If
                Next r
            Else
                reservaArray = CrearReservaArrayV12(arrRes, i, room, dateMapping)
                If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
            End If
        End If
NextI:
    Next i
    Set ProcesarReservasEnMemoriaV13 = reservasProcessadas
End Function

Private Function ProcesarBloqueosEnMemoriaV13(wsBloq As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    Dim bloqueosProcessados As Collection
    Set bloqueosProcessados = New Collection
    On Error Resume Next
    
    If wsBloq Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    Dim arrBloq As Variant
    arrBloq = wsBloq.Range("A2:C" & lastRow).Value
    If Not IsArray(arrBloq) Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    Dim i As Long, bloqueoArray As Variant, hoy As Date
    hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro
    If fechaCorte > 0 Then
        fechaFiltro = fechaCorte
    ElseIf soloFuturas Then
        fechaFiltro = hoy - 3
    Else
        fechaFiltro = 0
    End If
    
    For i = 1 To UBound(arrBloq, 1)
        If arrBloq(i, 1) <> "" And IsDate(arrBloq(i, 2)) And IsDate(arrBloq(i, 3)) Then
            If fechaFiltro > 0 Then
                If CDate(arrBloq(i, 3)) < fechaFiltro Then GoTo NextJ
            End If
            bloqueoArray = CrearBloqueoArrayV12(arrBloq, i, dateMapping)
            If bloqueoArray(0) > 0 Then bloqueosProcessados.Add bloqueoArray
        End If
NextJ:
    Next i
    Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados
End Function

'=================================================================================
' CREAR ARRAYS DE RESERVA Y BLOQUEO
'=================================================================================

Private Function CrearReservaArrayV12(arrRes As Variant, i As Long, ByVal room As String, dateMapping As Object) As Variant
    Dim reservaArray(0 To 3) As Variant
    reservaArray(0) = 0: reservaArray(1) = 0: reservaArray(2) = 0: reservaArray(3) = ""
    On Error Resume Next
    
    Select Case room
        Case "1": reservaArray(0) = 2
        Case "2": reservaArray(0) = 4
        Case "3": reservaArray(0) = 6
        Case "4": reservaArray(0) = 8
        Case "5": reservaArray(0) = 10
        Case "6": reservaArray(0) = 12
        Case "7": reservaArray(0) = 14
        Case "Of.1": reservaArray(0) = 16
        Case "Of.2": reservaArray(0) = 18
        Case "Of.3": reservaArray(0) = 20
        Case "Est.1": reservaArray(0) = 22
        Case "Est.2": reservaArray(0) = 24
        Case "Est.3": reservaArray(0) = 26
        Case Else: CrearReservaArrayV12 = reservaArray: Exit Function
    End Select
    
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario GIJ�N")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(reservaArray(0)).Hidden Then
            reservaArray(0) = 0
            CrearReservaArrayV12 = reservaArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    If Not IsDate(arrRes(i, 12)) Or Not IsDate(arrRes(i, 13)) Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    Dim checkin As Date, checkout As Date
    checkin = CDate(arrRes(i, 12))
    checkout = CDate(arrRes(i, 13))
    
    If dateMapping.Exists(CLng(checkin)) Then reservaArray(1) = dateMapping(CLng(checkin))
    If dateMapping.Exists(CLng(checkout)) Then reservaArray(2) = dateMapping(CLng(checkout))
    If reservaArray(1) = 0 Or reservaArray(2) = 0 Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    Dim fullName As String, displayName As String
    fullName = Trim(CStr(arrRes(i, 11)))
    
    If fullName = "" Then
        reservaArray(3) = arrRes(i, 1)
    Else
        ' Leer N� pax de columna O (posici?n 15 en el array, igual que en Soto)
    Dim pax As String
    pax = Trim(CStr(arrRes(i, 15)))
    
    ' Formatear nombre con pax entre par?ntesis
    ' L?mite de 15 caracteres para dejar espacio al "(N pax)"
    If Len(fullName) <= 15 Then
        displayName = fullName & " (" & pax & ")"
    Else
        displayName = Left(fullName, 15) & ". (" & pax & ")"
    End If
    reservaArray(3) = displayName & vbNewLine & arrRes(i, 1)
    End If
    
    If arrRes(i, 4) = "Comisi�n NO indem." Or arrRes(i, 4) = "Destino" Or arrRes(i, 4) = "Comisi�n" Then
        reservaArray(3) = reservaArray(3) & " C"
    End If
    CrearReservaArrayV12 = reservaArray
End Function

Private Function CrearBloqueoArrayV12(arrBloq As Variant, i As Long, dateMapping As Object) As Variant
    Dim bloqueoArray(0 To 3) As Variant
    bloqueoArray(0) = 0: bloqueoArray(1) = 0: bloqueoArray(2) = 0
    bloqueoArray(3) = CStr(arrBloq(i, 1))
    On Error Resume Next
    
    Select Case bloqueoArray(3)
        Case "1": bloqueoArray(0) = 2
        Case "2": bloqueoArray(0) = 4
        Case "3": bloqueoArray(0) = 6
        Case "4": bloqueoArray(0) = 8
        Case "5": bloqueoArray(0) = 10
        Case "6": bloqueoArray(0) = 12
        Case "7": bloqueoArray(0) = 14
        Case "Of.1": bloqueoArray(0) = 16
        Case "Of.2": bloqueoArray(0) = 18
        Case "Of.3": bloqueoArray(0) = 20
        Case "Est.1": bloqueoArray(0) = 22
        Case "Est.2": bloqueoArray(0) = 24
        Case "Est.3": bloqueoArray(0) = 26
        Case Else: CrearBloqueoArrayV12 = bloqueoArray: Exit Function
    End Select
    
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario GIJ�N")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(bloqueoArray(0)).Hidden Then
            bloqueoArray(0) = 0
            CrearBloqueoArrayV12 = bloqueoArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    If Not IsDate(arrBloq(i, 2)) Or Not IsDate(arrBloq(i, 3)) Then
        bloqueoArray(0) = 0: CrearBloqueoArrayV12 = bloqueoArray: Exit Function
    End If
    
    Dim dIni As Date, dFin As Date
    dIni = CDate(arrBloq(i, 2))
    dFin = CDate(arrBloq(i, 3))
    
    If dateMapping.Exists(CLng(dIni)) Then bloqueoArray(1) = dateMapping(CLng(dIni))
    If dateMapping.Exists(CLng(dFin)) Then bloqueoArray(2) = dateMapping(CLng(dFin))
    If bloqueoArray(1) = 0 Or bloqueoArray(2) = 0 Then bloqueoArray(0) = 0
    CrearBloqueoArrayV12 = bloqueoArray
End Function

'=================================================================================
' APLICAR RESERVAS Y BLOQUEOS
'=================================================================================

Private Function AplicarReservasYBloqueosConTramosV16(wsCal As Worksheet, reservas As Collection, bloqueos As Collection, mapaOcupacion As Object, solapeRecolector As Object) As Boolean
    Dim huboSolape As Boolean: huboSolape = False
    Dim i As Long, reservaArray As Variant, col As Long, clave As String
    Dim tramo As Long, bloqueoArray As Variant
    Dim colCheckin As Long, colCheckout As Long
    
    On Error Resume Next

    For i = 1 To reservas.Count
        reservaArray = reservas(i)
        
        colCheckin = reservaArray(1)
        colCheckout = reservaArray(2)
        
        ' DIA DE CHECK-IN - Solo TRAMO 3 (columna +2)
        clave = reservaArray(0) & "_" & (colCheckin + 2)
        With wsCal.Cells(reservaArray(0), colCheckin + 2)
            If .MergeCells Then .MergeCells = False
            If mapaOcupacion.Exists(clave) Then
            If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), colCheckin + 2) Then
                .Interior.color = COLOR_SOLAPE_ACEPTADO
            ' No marcamos huboSolape ni registramos
            Else
                .Interior.color = vbRed: huboSolape = True
                RegistrarSolape wsCal, colCheckin + 2, reservaArray, mapaOcupacion(clave), solapeRecolector
            End If
        Else
    .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)
End If
            .Value = reservaArray(3)
            .Font.Size = 15
            .Font.Bold = False
            Dim lineBreakPos As Long: lineBreakPos = InStr(.Value, vbNewLine)
            If lineBreakPos > 0 Then
                With .Characters(Start:=lineBreakPos + 1, Length:=Len(.Value) - lineBreakPos).Font
                    .Size = 30
                    .Bold = True
                End With
            End If
        End With
        
        ' DIAS INTERMEDIOS - TODOS los tramos (0, 1, 2)
        If colCheckout > colCheckin + 3 Then
            For col = colCheckin + 3 To colCheckout - 3 Step 3
                For tramo = 0 To 2
                    clave = reservaArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(reservaArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        If mapaOcupacion.Exists(clave) Then
    If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), col + tramo) Then
        .Interior.color = COLOR_SOLAPE_ACEPTADO
    Else
        .Interior.color = vbRed: huboSolape = True
        RegistrarSolape wsCal, col + tramo, reservaArray, mapaOcupacion(clave), solapeRecolector
    End If
Else
    .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)
End If
                    End With
                Next tramo
            Next col
        End If
        
        ' DIA DE CHECK-OUT - Solo TRAMO 1 (columna +0)
        If colCheckout > colCheckin Then
            clave = reservaArray(0) & "_" & colCheckout
            With wsCal.Cells(reservaArray(0), colCheckout)
                If .MergeCells Then .MergeCells = False
                If mapaOcupacion.Exists(clave) Then
    If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), colCheckout) Then
        .Interior.color = COLOR_SOLAPE_ACEPTADO
    Else
        .Interior.color = vbRed: huboSolape = True
        RegistrarSolape wsCal, colCheckout, reservaArray, mapaOcupacion(clave), solapeRecolector
    End If
Else
    .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)
End If
            End With
        End If
        
    Next i

    ' BLOQUEOS: pintado identico a reservas (intervalo semiabierto [inicio, fin))
    ' - Dia inicio: solo tramo checkin (col+2)
    ' - Dias intermedios: los 3 tramos
    ' - Dia fin: solo tramo checkout (col+0), si es distinto al dia inicio
    For i = 1 To bloqueos.Count
        bloqueoArray = bloqueos(i)
        
        ' DIA DE INICIO DEL BLOQUEO - solo tramo checkin (col+2)
        clave = bloqueoArray(0) & "_" & (bloqueoArray(1) + 2)
        With wsCal.Cells(bloqueoArray(0), bloqueoArray(1) + 2)
            If .MergeCells Then .MergeCells = False
            If mapaOcupacion.Exists(clave) Then
                If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), bloqueoArray(1) + 2) Then
                    .Interior.color = COLOR_SOLAPE_ACEPTADO
                Else
                    .Interior.color = vbRed: huboSolape = True
                    Dim numOrdenSolape As String
                    numOrdenSolape = "-"
                    If mapaOcupacion(clave) <> "BLOQUEO" Then
                        numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                    End If
                    RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(1) + 2, numOrdenSolape, bloqueoArray, solapeRecolector
                End If
            Else
                .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
            End If
            .Value = "BLOQUEADA"
        End With
        
        ' DIAS INTERMEDIOS DEL BLOQUEO - los 3 tramos
        If bloqueoArray(2) > bloqueoArray(1) + 3 Then
            For col = bloqueoArray(1) + 3 To bloqueoArray(2) - 3 Step 3
                For tramo = 0 To 2
                    clave = bloqueoArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(bloqueoArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        If mapaOcupacion.Exists(clave) Then
                            If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), col + tramo) Then
                                .Interior.color = COLOR_SOLAPE_ACEPTADO
                            Else
                                .Interior.color = vbRed: huboSolape = True
                                numOrdenSolape = "-"
                                If mapaOcupacion(clave) <> "BLOQUEO" Then
                                    numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                                End If
                                RegistrarSolapeReservaBloqueo wsCal, col + tramo, numOrdenSolape, bloqueoArray, solapeRecolector
                            End If
                        Else
                            .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
                        End If
                    End With
                Next tramo
                wsCal.Cells(bloqueoArray(0), col + 2).Value = "BLOQUEADA"
            Next col
        End If
        
        ' DIA DE FIN DEL BLOQUEO - solo tramo checkout (col+0), si es distinto al inicio
        If bloqueoArray(2) > bloqueoArray(1) Then
            clave = bloqueoArray(0) & "_" & bloqueoArray(2)
            With wsCal.Cells(bloqueoArray(0), bloqueoArray(2))
                If .MergeCells Then .MergeCells = False
                If mapaOcupacion.Exists(clave) Then
                    If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), bloqueoArray(2)) Then
                        .Interior.color = COLOR_SOLAPE_ACEPTADO
                    Else
                        .Interior.color = vbRed: huboSolape = True
                        numOrdenSolape = "-"
                        If mapaOcupacion(clave) <> "BLOQUEO" Then
                            numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                        End If
                        RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(2), numOrdenSolape, bloqueoArray, solapeRecolector
                    End If
                Else
                    .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
                End If
            End With
        End If
    Next i
    
    AplicarReservasYBloqueosConTramosV16 = huboSolape
End Function

'=================================================================================
' REPINTAR RESERVAS PAGADAS
'=================================================================================

Private Sub RepintarReservasPagadas(wsRes As Worksheet, wsCal As Worksheet, reservasProcessadas As Collection)
    On Error Resume Next
    
    If reservasProcessadas Is Nothing Then Exit Sub
    If reservasProcessadas.Count = 0 Then Exit Sub
    
    Dim i As Long, pagado As String, filaRes As Long, reservaArray As Variant
    Dim col As Long, tramo As Long, colCheckin As Long, colCheckout As Long
    
    For i = 1 To reservasProcessadas.Count
        reservaArray = reservasProcessadas(i)
        Dim partes() As String
        
        If InStr(reservaArray(3), vbNewLine) > 0 Then
            partes = Split(reservaArray(3), vbNewLine)
            If UBound(partes) > 0 Then
                Dim numOrden As Variant
                numOrden = Trim(Split(partes(1), " ")(0))
                filaRes = BuscarEnColumna(wsRes, 1, numOrden)
                If filaRes > 0 Then
                    pagado = UCase(wsRes.Cells(filaRes, 27).Value)
                    If MarcarSiPagadosEnResidencia.EsEstadoPagado(pagado) Then
                        colCheckin = reservaArray(1)
                        colCheckout = reservaArray(2)
                        
                        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = RGB(198, 239, 206)
                        
                        If colCheckout > colCheckin + 3 Then
                            For col = colCheckin + 3 To colCheckout - 3 Step 3
                                For tramo = 0 To 2
                                    wsCal.Cells(reservaArray(0), col + tramo).Interior.color = RGB(198, 239, 206)
                                Next tramo
                            Next col
                        End If
                        
                        If colCheckout > colCheckin Then
                            wsCal.Cells(reservaArray(0), colCheckout).Interior.color = RGB(198, 239, 206)
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

'=================================================================================
' BUSQUEDAS Y UTILIDADES
'=================================================================================

Private Function BuscarEnColumna(ws As Worksheet, numCol As Long, valor As Variant) As Long
    On Error Resume Next
    BuscarEnColumna = 0
    If ws Is Nothing Then Exit Function
    If numCol < 1 Or numCol > ws.Columns.Count Then Exit Function
    
    Dim ultimaFila As Long, i As Long
    ultimaFila = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row
    If ultimaFila < 2 Then Exit Function
    
    For i = 2 To ultimaFila
        If Trim(CStr(ws.Cells(i, numCol).Value)) = Trim(CStr(valor)) Then
            BuscarEnColumna = i
            Exit Function
        End If
    Next i
End Function

'=================================================================================
' REGISTRO DE SOLAPES
'=================================================================================

Private Sub RegistrarSolape(wsCal As Worksheet, ByVal colIndex As Long, arr1 As Variant, arr2 As Variant, solapeRecolector As Object, Optional esBloqueo As Boolean = False)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, numOrdenReserva As String, numOrdenReserva2 As String
    Dim claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    ' CORREGIDO: Obtener habStr ANTES de usarlo en el filtro
    habStr = ObtenerNumHab(arr1(0))
    
    ' --- FILTRO: SOLAPE ACEPTADO POR GRUPO (HAB + RANGO) ---
    Dim claveGrupo As String
    claveGrupo = ClaveSolapeGrupoGijon(habStr)

    If Not solapesAceptados Is Nothing Then
        If SolapeAceptadoGijon(solapesAceptados, claveGrupo, CDate(fechaLng)) Then
            Exit Sub
        End If
    End If
    
    numOrdenReserva = ObtenerNumOrden(arr1(3))
    numOrdenReserva2 = ObtenerNumOrden(arr2)
    
    If numOrdenReserva = "-" Or numOrdenReserva2 = "-" Then Exit Sub
    
    Dim ord1 As Long, ord2 As Long
    If Not IsNumeric(numOrdenReserva) Or Not IsNumeric(numOrdenReserva2) Then Exit Sub
    
    ord1 = CLng(numOrdenReserva)
    ord2 = CLng(numOrdenReserva2)
    
    If ord1 > ord2 Then
        Dim temp As Long
        temp = ord1: ord1 = ord2: ord2 = temp
    End If
    
    claveSolape = "HAB_" & habStr & "_ORD_" & CStr(ord1) & "_" & CStr(ord2)
    
    If Not solapeRecolector.Exists(claveSolape) Then
        solapeRecolector.Add claveSolape, habStr & "|" & CStr(ord1) & "|" & CStr(ord2) & "|" & CStr(fechaLng) & ","
    Else
        solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
    End If
End Sub

Private Sub RegistrarSolapeReservaBloqueo(wsCal As Worksheet, ByVal colIndex As Long, numOrdenReserva As String, bloqueoArr As Variant, solapeRecolector As Object)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If Trim(CStr(numOrdenReserva)) = "-" Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    habStr = ObtenerNumHab(bloqueoArr(0))
    
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    If numOrdenReserva <> "-" And IsNumeric(numOrdenReserva) Then
        claveSolape = "HAB_" & habStr & "_ORD_" & numOrdenReserva & "_BLOQUEO"
        If Not solapeRecolector.Exists(claveSolape) Then
            solapeRecolector.Add claveSolape, habStr & "|" & numOrdenReserva & "|BLOQUEO|" & CStr(fechaLng) & ","
        Else
            solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
        End If
    End If
End Sub

Private Function EsSolapeGrupoAceptadoEnCelda(ByVal wsCal As Worksheet, ByVal filaHab As Long, ByVal colIndex As Long) As Boolean
    On Error Resume Next
    EsSolapeGrupoAceptadoEnCelda = False
    If wsCal Is Nothing Then Exit Function
    If solapesAceptados Is Nothing Then Exit Function
    If filaHab < 1 Or filaHab > wsCal.Rows.Count Then Exit Function
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Function

    ' Calcular columna "inicio del d?a" (tus d�as van de 3 en 3, empezando en col 2)
    Dim dayFirstCol As Long
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2

    Dim fechaCelda As Variant
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Function

    Dim habStr As String
    habStr = ObtenerNumHab(filaHab)
    If habStr = "?" Or Trim$(habStr) = "" Then Exit Function

    Dim claveGrupo As String
    claveGrupo = ClaveSolapeGrupoGijon(habStr)

    EsSolapeGrupoAceptadoEnCelda = SolapeAceptadoGijon(solapesAceptados, claveGrupo, CDate(fechaCelda))
End Function

'=================================================================================
' UTILIDADES PARA TEXTOS
'=================================================================================

Private Function ObtenerNumHab(ByVal filaHab As Variant) As String
    On Error Resume Next
    Select Case CLng(filaHab)
        Case 2: ObtenerNumHab = "1"
        Case 4: ObtenerNumHab = "2"
        Case 6: ObtenerNumHab = "3"
        Case 8: ObtenerNumHab = "4"
        Case 10: ObtenerNumHab = "5"
        Case 12: ObtenerNumHab = "6"
        Case 14: ObtenerNumHab = "7"
        Case 16: ObtenerNumHab = "Of.1"
        Case 18: ObtenerNumHab = "Of.2"
        Case 20: ObtenerNumHab = "Of.3"
        Case 22: ObtenerNumHab = "Est.1"
        Case 24: ObtenerNumHab = "Est.2"
        Case 26: ObtenerNumHab = "Est.3"
        Case Else: ObtenerNumHab = "?"
    End Select
End Function

Private Function ObtenerNumOrden(ByVal textoReserva As Variant) As String
    On Error Resume Next
    ObtenerNumOrden = "-"
    
    If IsNull(textoReserva) Then Exit Function
    Dim s As String
    s = CStr(textoReserva)
    If s = "" Or s = "BLOQUEO" Then Exit Function
    
    If InStr(s, vbNewLine) = 0 Then
        s = Trim(Replace(s, " C", ""))
        If IsNumeric(s) Then ObtenerNumOrden = s
        Exit Function
    End If
    
    Dim partes() As String
    partes = Split(s, vbNewLine)
    If UBound(partes) < 1 Then Exit Function
    
    Dim lineaNumero As String
    lineaNumero = Trim(partes(1))
    
    If Right(lineaNumero, 2) = " C" Then
        lineaNumero = Left(lineaNumero, Len(lineaNumero) - 2)
    End If
    
    Dim posibleNum As String
    posibleNum = Trim(Split(lineaNumero, " ")(0))
    If IsNumeric(posibleNum) Then ObtenerNumOrden = posibleNum
End Function

'=================================================================================
' MOSTRAR DETALLE DE SOLAPES
'=================================================================================

Private Sub MostrarDetalleSolapes()
    On Error Resume Next
    If tipoSolape Is Nothing Then Exit Sub
    If tipoSolape.Count = 0 Then
        MsgBox "No hay solapes registrados.", vbInformation, "Sin Solapes"
        Exit Sub
    End If
    
    Dim detalles As String
    detalles = "ATENCION: Solapes Detectados" & vbCrLf & vbCrLf & _
               "Hay reservas que solapan entre si o con bloqueos." & vbCrLf & _
               "Las celdas afectadas aparecen en ROJO." & vbCrLf & vbCrLf & _
               "DETALLE DE SOLAPES DETECTADOS:" & vbCrLf & vbCrLf
    
    Dim k As Variant, info As String, partes() As String
    Dim habStr As String, ord1 As String, ord2 As String, fechasStr As String
    Dim fechasArray() As String, fechasLong() As Long, i As Long, rangoTexto As String
    
    For Each k In tipoSolape.Keys
        info = CStr(tipoSolape(k))
        If InStr(info, "|") = 0 Then GoTo NextSolape
        
        partes = Split(info, "|")
        If UBound(partes) < 3 Then GoTo NextSolape
        
        habStr = Trim(partes(0))
        ord1 = Trim(partes(1))
        ord2 = Trim(partes(2))
        fechasStr = Trim(partes(3))
        
        If Right(fechasStr, 1) = "," Then fechasStr = Left(fechasStr, Len(fechasStr) - 1)
        If fechasStr = "" Then GoTo NextSolape
        
        If InStr(fechasStr, ",") > 0 Then
            fechasArray = Split(fechasStr, ",")
        Else
            ReDim fechasArray(0 To 0)
            fechasArray(0) = fechasStr
        End If
        
        ReDim fechasLong(LBound(fechasArray) To UBound(fechasArray))
        For i = LBound(fechasArray) To UBound(fechasArray)
            If IsNumeric(Trim(fechasArray(i))) Then
                fechasLong(i) = CLng(Trim(fechasArray(i)))
            End If
        Next i
        
        rangoTexto = AgruparFechasEnRangos(fechasLong)
        
        Dim padding As String
        If Len(rangoTexto) < 18 Then
            padding = Space$(18 - Len(rangoTexto))
        Else
            padding = " "
        End If
        If ord2 = "BLOQUEO" Then
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " Hab. " & habStr & vbTab & Chr(187) & " [!] Reserva N" & Chr(186) & " " & ord1 & " con BLOQUEO" & vbCrLf
        Else
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " Hab. " & habStr & vbTab & Chr(187) & " [!] Reservas N" & Chr(186) & " " & ord1 & " y " & ord2 & vbCrLf
        End If
        
NextSolape:
    Next k
    
    MsgBox detalles, vbExclamation, "Solapes Detectados"
End Sub

Private Function AgruparFechasEnRangos(fechasLong() As Long) As String
    On Error Resume Next
    
    If UBound(fechasLong) < LBound(fechasLong) Then
        AgruparFechasEnRangos = "Sin fechas"
        Exit Function
    End If
    
    Dim fechasList As Collection
    Set fechasList = New Collection
    Dim i As Long
    
    For i = LBound(fechasLong) To UBound(fechasLong)
        If fechasLong(i) > 0 Then
            InsertarFechaOrdenada fechasList, CDate(fechasLong(i))
        End If
    Next i
    
    If fechasList.Count = 0 Then
        AgruparFechasEnRangos = "Sin fechas"
        Exit Function
    End If
    
    Dim resultado As String
    Dim inicioRango As Date, finRango As Date
    Dim primerRango As Boolean: primerRango = True
    
    inicioRango = fechasList(1)
    finRango = fechasList(1)
    
    For i = 2 To fechasList.Count
        If fechasList(i) = finRango + 1 Then
            finRango = fechasList(i)
        Else
            If Not primerRango Then resultado = resultado & ", "
            primerRango = False
            
            If inicioRango = finRango Then
                resultado = resultado & FormatearFechaSolape(inicioRango)
            Else
                resultado = resultado & FormatearFechaSolape(inicioRango) & " al " & FormatearFechaSolape(finRango)
            End If
            
            inicioRango = fechasList(i)
            finRango = fechasList(i)
        End If
    Next i
    
    If Not primerRango Then resultado = resultado & ", "
    
    If inicioRango = finRango Then
        resultado = resultado & FormatearFechaSolape(inicioRango)
    Else
        resultado = resultado & FormatearFechaSolape(inicioRango) & " al " & FormatearFechaSolape(finRango)
    End If
    
    If Year(inicioRango) <> Year(Date) Then
        resultado = resultado & "-" & Year(inicioRango)
    End If
    
    AgruparFechasEnRangos = resultado
End Function

Private Function FormatearFechaSolape(ByVal d As Date) As String
    Dim s As String
    s = Format(d, "d-mmm")
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    
    Dim partes() As String
    partes = Split(s, "-")
    If UBound(partes) = 1 Then
        Dim mes As String
        mes = partes(1)
        If Len(mes) > 0 Then
            mes = UCase$(Left$(mes, 1)) & Mid$(mes, 2)
        End If
        FormatearFechaSolape = partes(0) & "-" & mes
    Else
        FormatearFechaSolape = s
    End If
End Function

Private Sub InsertarFechaOrdenada(ByRef col As Collection, fecha As Date)
    On Error Resume Next
    
    If col.Count = 0 Then
        col.Add fecha
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To col.Count
        If fecha < col(i) Then
            col.Add fecha, Before:=i
            Exit Sub
        ElseIf fecha = col(i) Then
            Exit Sub
        End If
    Next i
    
    col.Add fecha
End Sub

'=================================================================================
' CONTAR RESERVAS EN HABITACIONES OCULTAS
'=================================================================================

Private Function ContarReservasEnHabitacionesOcultas(wsRes As Worksheet, wsCal As Worksheet, ByRef listaOcultas As String) As Long
    ' OBJETIVO: Contar reservas actuales/futuras asignadas a habitaciones
    '           que est�n ocultos/as en el calendario (filas ocultas)
    '           y recopilar una lista de sus nombres
    
    On Error Resume Next
    
    ContarReservasEnHabitacionesOcultas = 0
    
    If wsRes Is Nothing Or wsCal Is Nothing Then Exit Function
    
    Dim lastRow As Long, i As Long
    Dim fechaSalida As Date, hoy As Date
    Dim habitaciones As String, arrHabs As Variant
    Dim hab As Variant, filaHab As Long
    Dim contadorOcultas As Long
    
    hoy = Date
    contadorOcultas = 0
    
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    For i = 2 To lastRow
        If IsDate(wsRes.Cells(i, "M").Value) And Not IsEmpty(wsRes.Cells(i, "S").Value) Then
            fechaSalida = wsRes.Cells(i, "M").Value
            
            If fechaSalida >= hoy Then
                habitaciones = Trim(CStr(wsRes.Cells(i, "S").Value))
                
                If InStr(habitaciones, ",") > 0 Then
                    arrHabs = Split(habitaciones, ",")
                    Dim resContada As Boolean: resContada = False
                    For Each hab In arrHabs
                        filaHab = MapearHabitacionAFila(Trim(CStr(hab)))
                        If filaHab > 0 Then
                            If wsCal.Rows(filaHab).Hidden Then
                                If Not resContada Then
                                    contadorOcultas = contadorOcultas + 1
                                    resContada = True
                                End If
                                Dim rStr As String
                                rStr = Trim(CStr(hab))
                                If InStr("," & listaOcultas & ",", "," & rStr & ",") = 0 Then
                                    If listaOcultas = "" Then
                                        listaOcultas = rStr
                                    Else
                                        listaOcultas = listaOcultas & ", " & rStr
                                    End If
                                End If
                            End If
                        End If
                    Next hab
                Else
                    filaHab = MapearHabitacionAFila(habitaciones)
                    If filaHab > 0 Then
                        If wsCal.Rows(filaHab).Hidden Then
                            contadorOcultas = contadorOcultas + 1
                            Dim rStr2 As String
                            rStr2 = Trim(CStr(habitaciones))
                            If InStr("," & listaOcultas & ",", "," & rStr2 & ",") = 0 Then
                                If listaOcultas = "" Then
                                    listaOcultas = rStr2
                                Else
                                    listaOcultas = listaOcultas & ", " & rStr2
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    ContarReservasEnHabitacionesOcultas = contadorOcultas
End Function

Private Function MapearHabitacionAFila(ByVal habitacion As String) As Long
    On Error Resume Next
    MapearHabitacionAFila = 0
    
    Select Case Trim(habitacion)
        Case "1": MapearHabitacionAFila = 2
        Case "2": MapearHabitacionAFila = 4
        Case "3": MapearHabitacionAFila = 6
        Case "4": MapearHabitacionAFila = 8
        Case "5": MapearHabitacionAFila = 10
        Case "6": MapearHabitacionAFila = 12
        Case "7": MapearHabitacionAFila = 14
        Case "Of.1": MapearHabitacionAFila = 16
        Case "Of.2": MapearHabitacionAFila = 18
        Case "Of.3": MapearHabitacionAFila = 20
        Case "Est.1": MapearHabitacionAFila = 22
        Case "Est.2": MapearHabitacionAFila = 24
        Case "Est.3": MapearHabitacionAFila = 26
    End Select
End Function



