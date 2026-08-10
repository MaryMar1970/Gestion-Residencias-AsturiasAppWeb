VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeleccionHabitacionOVIEDO 
   Caption         =   "Selecciona Habitaci�n OVIEDO"
   ClientHeight    =   5055
   ClientLeft      =   300
   ClientTop       =   1110
   ClientWidth     =   3255
   OleObjectBlob   =   "SeleccionHabitacionOVIEDO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SeleccionHabitacionOVIEDO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Public habitacionesDisponibles As String
Public filaActual As Long

Private Sub UserForm_Initialize()
    Dim disponiblesString As String
    Dim disponibles() As String
    Dim listaHabitacionesVisible As Collection
    Dim i As Long
    Dim habitacion As String
    Dim wsCal As Worksheet, wsRes As Worksheet, wsBloq As Worksheet
    Dim filaHab As Long, filaActual As Long, filaBloq As Long
    Dim fechaEntrada As Date, fechaSalida As Date
    Dim arrBloq As Variant, j As Long, bloqueada As Boolean
    Dim desdeBloq As Date, hastaBloq As Date
    Dim paxValue As Long

    On Error GoTo SinDisponibles

    Application.enableEvents = False

    disponiblesString = ThisWorkbook.Names("HabitacionesDisponiblesOVIEDO").RefersTo
    disponiblesString = Replace(disponiblesString, "=", "")
    disponiblesString = Replace(disponiblesString, Chr(34), "")
    disponibles = Split(disponiblesString, ",")

    Set wsCal = ThisWorkbook.Worksheets("Calendario OVIEDO")
    Set wsRes = ThisWorkbook.Sheets("RESIDENCIA OVIEDO")
    On Error Resume Next
    Set wsBloq = ThisWorkbook.Sheets("Habitaciones Bloqueadas OVIEDO")
    On Error GoTo SinDisponibles

    filaActual = ActiveCell.Row
    fechaEntrada = wsRes.Cells(filaActual, "L").Value
    fechaSalida = wsRes.Cells(filaActual, "M").Value
    paxValue = val(wsRes.Cells(filaActual, "O").Value)

    ' Cargar bloqueos
    arrBloq = Empty
    If Not wsBloq Is Nothing Then
        On Error Resume Next
        filaBloq = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
        If Err.Number = 0 And filaBloq >= 2 Then
            arrBloq = wsBloq.Range("A2:D" & filaBloq).Value
        End If
        On Error GoTo SinDisponibles
    End If

    Set listaHabitacionesVisible = New Collection
    For i = 0 To UBound(disponibles)
        habitacion = Trim(disponibles(i))
        filaHab = ObtenerFilaDeHabitacion(habitacion)
        If filaHab > 0 Then
            If Not wsCal.Rows(filaHab).Hidden Then
                listaHabitacionesVisible.Add habitacion
            End If
        End If
    Next i

    Me.ListBox1.Clear
    Me.ListBox1.MultiSelect = fmMultiSelectMulti

    For i = 1 To listaHabitacionesVisible.Count
        habitacion = listaHabitacionesVisible(i)
        bloqueada = False

        ' Comprobar solape con bloqueos
        If Not IsEmpty(arrBloq) And IsDate(fechaEntrada) And IsDate(fechaSalida) Then
            For j = 1 To UBound(arrBloq, 1)
                If StrComp(Trim(habitacion), Trim(arrBloq(j, 1)), vbTextCompare) = 0 Then
                    If IsDate(arrBloq(j, 2)) And IsDate(arrBloq(j, 3)) Then
                        desdeBloq = arrBloq(j, 2)
                        hastaBloq = arrBloq(j, 3)
                            If Not (fechaSalida <= desdeBloq Or fechaEntrada >= IIf(desdeBloq = hastaBloq, hastaBloq + 1, hastaBloq)) Then
                            bloqueada = True
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If

        If bloqueada Then
            Me.ListBox1.AddItem habitacion & " - BLOQ."
        Else
            Me.ListBox1.AddItem habitacion
        End If
    Next i

    Me.AceptarButton.SetFocus

    Application.enableEvents = True
    Exit Sub

SinDisponibles:
    Application.enableEvents = True
    MsgBox "No se pudo recuperar la lista de habitaciones disponibles.", vbExclamation, "Error"
    Unload Me
End Sub

Private Function ObtenerFilaDeHabitacion(ByVal habitacion As String) As Long
    ObtenerFilaDeHabitacion = ObtenerFilaDeHabitacionOviedo(habitacion)
End Function

Private Sub AceptarButton_Click()
    On Error GoTo ManejoErrores

    Dim i As Long, seleccionadas As String
    Dim seleccionArray() As String
    Dim wsBloq As Worksheet, wsRes As Worksheet
    Dim fechaEntrada As Date, fechaSalida As Date
    Dim filaActual As Long
    Dim filaBloq As Long, arrBloq As Variant
    Dim j As Long
    Dim haySolape As Boolean
    Dim mensajeSolape As String

    Dim itemStr As String
    seleccionadas = ""
    For i = 0 To Me.ListBox1.ListCount - 1
        If Me.ListBox1.Selected(i) Then
            itemStr = Me.ListBox1.List(i)
            If InStr(itemStr, " - ") > 0 Then
                itemStr = Split(itemStr, " - ")(0)
            End If
            If seleccionadas <> "" Then seleccionadas = seleccionadas & ", "
            seleccionadas = seleccionadas & itemStr
        End If
    Next i

    If seleccionadas = "" Then
        MsgBox "Debes seleccionar al menos una habitaci�n.", vbExclamation
        Exit Sub
    End If

    Set wsBloq = ThisWorkbook.Sheets("Habitaciones Bloqueadas OVIEDO")
    Set wsRes = ThisWorkbook.Sheets("RESIDENCIA OVIEDO")
    
    filaActual = ActiveCell.Row

    ' --- NUEVO: Validar capacidad ---
    Dim paxVal As Long
    Dim camasSupletorias As Long
    Dim capacidadSel As Long
    Dim habTemp As Variant
    Dim arrTemp() As String
    
    paxVal = val(wsRes.Cells(filaActual, "O").Value)
    camasSupletorias = val(wsRes.Cells(filaActual, "S").Value) ' S = Camas Supletorias en Oviedo
    capacidadSel = camasSupletorias
    arrTemp = Split(seleccionadas, ",")
    
        For Each habTemp In arrTemp
        If Trim(CStr(habTemp)) <> "" Then
            capacidadSel = capacidadSel + ObtenerCapacidadOviedo(CStr(habTemp))
        End If
    Next habTemp
    
    If paxVal > capacidadSel Then
        Dim respSobre As VbMsgBoxResult
        respSobre = MsgBox("El n�mero de PAX (" & paxVal & ") supera la capacidad m�xima de las habitaciones seleccionadas (" & capacidadSel & ")." & vbCrLf & vbCrLf & _
                           "�Desea continuar con la adjudicaci�n?", vbExclamation + vbYesNo, "Aviso de sobreocupaci�n")
        If respSobre = vbNo Then
            Exit Sub
        End If
    End If
    
    If Not IsDate(wsRes.Cells(filaActual, "L").Value) Or Not IsDate(wsRes.Cells(filaActual, "M").Value) Then
        MsgBox "Las fechas de entrada y salida deben estar definidas antes de seleccionar habitaciones.", vbExclamation
        Exit Sub
    End If
    
    fechaEntrada = wsRes.Cells(filaActual, "L").Value
    fechaSalida = wsRes.Cells(filaActual, "M").Value

    If fechaSalida <= fechaEntrada Then
        MsgBox "La fecha de salida debe ser posterior a la fecha de entrada.", vbExclamation
        Exit Sub
    End If

    filaBloq = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    If filaBloq >= 2 Then
        arrBloq = wsBloq.Range("A2:D" & filaBloq).Value
    Else
        arrBloq = Empty
    End If

    seleccionArray = Split(seleccionadas, ",")
    haySolape = False
    mensajeSolape = ""

    If Not IsEmpty(arrBloq) Then
        For i = 0 To UBound(seleccionArray)
            For j = 1 To UBound(arrBloq, 1)
                If StrComp(Trim(seleccionArray(i)), Trim(arrBloq(j, 1)), vbTextCompare) = 0 Then
                    If IsDate(arrBloq(j, 2)) And IsDate(arrBloq(j, 3)) Then
                        Dim desdeBloq As Date, hastaBloq As Date
                        desdeBloq = arrBloq(j, 2)
                        hastaBloq = arrBloq(j, 3)
                        
                            If Not (fechaSalida <= desdeBloq Or fechaEntrada >= IIf(desdeBloq = hastaBloq, hastaBloq + 1, hastaBloq)) Then
                            haySolape = True
                            mensajeSolape = mensajeSolape & vbCrLf & _
                                "Habitaci�n: " & Trim(seleccionArray(i)) & _
                                " | Bloqueada del " & Format(desdeBloq, "dd/mm/yyyy") & _
                                " al " & Format(hastaBloq, "dd/mm/yyyy") & _
                                IIf(Not IsEmpty(arrBloq(j, 4)), " | Motivo: " & arrBloq(j, 4), "")
                        End If
                    End If
                End If
            Next j
        Next i
    End If

If haySolape Then
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("�Atenci�n! La actual selecci�n se solapa parcialmente con fechas bloqueadas:" & _
                       vbCrLf & mensajeSolape & vbCrLf & vbCrLf & _
                       "�Desea continuar con la selecci�n y eliminar el bloqueo?", _
                       vbExclamation + vbYesNo, "Solape con bloqueo")
            
    If respuesta = vbNo Then
        Exit Sub
    Else
        ' Eliminar o reajustar los bloqueos en orden inverso para evitar problemas de indices de las filas
        For j = UBound(arrBloq, 1) To 1 Step -1
            For i = 0 To UBound(seleccionArray)
                If StrComp(Trim(seleccionArray(i)), Trim(arrBloq(j, 1)), vbTextCompare) = 0 Then
                    If IsDate(arrBloq(j, 2)) And IsDate(arrBloq(j, 3)) Then
                        desdeBloq = arrBloq(j, 2)
                        hastaBloq = arrBloq(j, 3)
                        If Not (fechaSalida <= desdeBloq Or fechaEntrada >= IIf(desdeBloq = hastaBloq, hastaBloq + 1, hastaBloq)) Then
                            
                            ' EVALUACI�N DE LOS 4 CASOS DE SOLAPE (Encoger, dividir o eliminar)
                            If fechaEntrada <= desdeBloq And fechaSalida >= hastaBloq Then
                                ' Caso 3: Solape total -> Eliminar el bloqueo completo
                                wsBloq.Rows(j + 1).Delete
                                
                            ElseIf fechaEntrada <= desdeBloq And fechaSalida < hastaBloq Then
                                ' Caso 1: Solape al inicio -> Encoger por la izquierda
                                wsBloq.Cells(j + 1, 2).Value = fechaSalida
                                
                            ElseIf fechaEntrada > desdeBloq And fechaSalida >= hastaBloq Then
                                ' Caso 2: Solape al final -> Encoger por la derecha
                                wsBloq.Cells(j + 1, 3).Value = fechaEntrada
                                
                            ElseIf fechaEntrada > desdeBloq And fechaSalida < hastaBloq Then
                                ' Caso 4: Solape interno -> Dividir el bloqueo en dos partes
                                Dim motivo As String
                                motivo = wsBloq.Cells(j + 1, 4).Value ' Columna D (Motivo)
                                
                                ' 1. Modificar el actual para que termine donde empieza la reserva
                                wsBloq.Cells(j + 1, 3).Value = fechaEntrada
                                
                                ' 2. Crear una nueva fila justo debajo para el tramo restante
                                wsBloq.Rows(j + 2).Insert Shift:=xlShiftDown
                                wsBloq.Cells(j + 2, 1).Value = arrBloq(j, 1) ' Habitaci�n
                                wsBloq.Cells(j + 2, 2).Value = fechaSalida     ' Inicia donde termina la reserva
                                wsBloq.Cells(j + 2, 3).Value = hastaBloq       ' Termina en la fecha fin original
                                If motivo <> "" Then wsBloq.Cells(j + 2, 4).Value = motivo
                            End If
                            
                        End If
                    End If
                End If
            Next i
        Next j
    End If
End If

    wsRes.Cells(filaActual, "T").Value = seleccionadas
    
    ' --- RESOLUCION automatica: CONCEDIDA o SI segun dias de antelacion ---
    If IsDate(wsRes.Cells(filaActual, "B").Value) And IsDate(wsRes.Cells(filaActual, "L").Value) Then
        Dim fechaSolO As Date, umbralO As Long, diasAntO As Long
        fechaSolO = CDate(wsRes.Cells(filaActual, "B").Value)
        umbralO = 14 - Weekday(fechaSolO, vbMonday)
        diasAntO = DateDiff("d", fechaSolO, CDate(wsRes.Cells(filaActual, "L").Value))
        Application.enableEvents = False
        If diasAntO <= umbralO Then
            wsRes.Cells(filaActual, "P").Value = "CONCEDIDA"
        Else
            wsRes.Cells(filaActual, "P").Value = "SI"
        End If
        Application.enableEvents = True
    End If
    
    Unload Me
    Exit Sub

ManejoErrores:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical
    Exit Sub
End Sub

Private Sub CancelarButton_Click()
    ' --- RESOLUCION automatica por cancelacion: DENEGADA o NO segun dias de antelacion ---
    Dim wsResCO As Worksheet
    Dim filaCO As Long
    Set wsResCO = ThisWorkbook.Sheets("RESIDENCIA OVIEDO")
    filaCO = ActiveCell.Row
    Dim valorActualCO As String
    valorActualCO = UCase(Trim(wsResCO.Cells(filaCO, "P").Value))
    If valorActualCO <> "SI" And valorActualCO <> "CONCEDIDA" And _
       valorActualCO <> "REEVALUADA" And valorActualCO <> "RENUNCIA" Then
        If IsDate(wsResCO.Cells(filaCO, "B").Value) And IsDate(wsResCO.Cells(filaCO, "L").Value) Then
            Dim fechaSolCO As Date, umbralCO As Long, diasAntCO As Long
            fechaSolCO = CDate(wsResCO.Cells(filaCO, "B").Value)
            umbralCO = 14 - Weekday(fechaSolCO, vbMonday)
            diasAntCO = DateDiff("d", fechaSolCO, CDate(wsResCO.Cells(filaCO, "L").Value))
            Application.enableEvents = False
            If diasAntCO <= umbralCO Then
                wsResCO.Cells(filaCO, "P").Value = "DENEGADA"
            Else
                wsResCO.Cells(filaCO, "P").Value = "NO"
            End If
            Application.enableEvents = True
        End If
    End If
    Unload Me
End Sub

