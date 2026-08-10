VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} SeleccionHabitacionSOTO 
   Caption         =   "Selecciona Apartamento SOTO"
   ClientHeight    =   4410
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   3255
   OleObjectBlob   =   "SeleccionHabitacionSOTO.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "SeleccionHabitacionSOTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
' Forzar la declaraci�n expl�cita de variables para evitar errores de escritura y mejorar el mantenimiento

Public habitacionesDisponibles As String
Public filaActual As Long

' Evento Click del ListBox1 (actualmente vac�o)
Private Sub ListBox1_Click()
    ' Este evento se activa al hacer clic en un elemento del ListBox
    ' No contiene implementaci�n en este momento
End Sub

' Evento Initialize del UserForm (se ejecuta al cargar el formulario)
Sub UserForm_Initialize()
    Dim disponiblesString As String
    Dim disponibles() As String
    Dim listaApartamentosVisible As Collection
    Dim i As Long
    Dim apartamento As String
    Dim wsCal As Worksheet, wsRes As Worksheet, wsBloq As Worksheet
    Dim filaApt As Long, filaActual As Long, filaBloq As Long
    Dim fechaEntrada As Date, fechaSalida As Date
    Dim arrBloq As Variant, j As Long, bloqueado As Boolean
    Dim desdeBloq As Date, hastaBloq As Date

    On Error GoTo SinDisponibles

    disponiblesString = ThisWorkbook.Names("ApartamentosDisponibles").RefersTo
    disponiblesString = Replace(disponiblesString, "=", "")
    disponiblesString = Replace(disponiblesString, """", "")
    disponibles = Split(disponiblesString, ",")

    Set wsCal = ThisWorkbook.Worksheets("Calendario SOTO")
    Set wsRes = ThisWorkbook.Sheets("RESIDENCIA SOTO")
    On Error Resume Next
    Set wsBloq = ThisWorkbook.Sheets("Apartamentos Bloqueados SOTO")
    On Error GoTo SinDisponibles

    filaActual = ActiveCell.Row
    fechaEntrada = wsRes.Cells(filaActual, "L").Value
    fechaSalida = wsRes.Cells(filaActual, "M").Value

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

    Set listaApartamentosVisible = New Collection
    
    For i = 0 To UBound(disponibles)
        apartamento = Trim(disponibles(i))
        filaApt = ObtenerFilaDeApartamento(apartamento)
        
        If filaApt > 0 Then
            If Not wsCal.Rows(filaApt).Hidden Then
                listaApartamentosVisible.Add apartamento
            End If
        End If
    Next i

    Me.ListBox1.Clear
    Me.ListBox1.MultiSelect = fmMultiSelectSingle
    
    For i = 1 To listaApartamentosVisible.Count
        apartamento = listaApartamentosVisible(i)
        bloqueado = False

        ' Comprobar solape con bloqueos
        If Not IsEmpty(arrBloq) And IsDate(fechaEntrada) And IsDate(fechaSalida) Then
            For j = 1 To UBound(arrBloq, 1)
                If StrComp(Trim(apartamento), Trim(arrBloq(j, 1)), vbTextCompare) = 0 Then
                    If IsDate(arrBloq(j, 2)) And IsDate(arrBloq(j, 3)) Then
                        desdeBloq = arrBloq(j, 2)
                        hastaBloq = arrBloq(j, 3)
                            If Not (fechaSalida <= desdeBloq Or fechaEntrada >= IIf(desdeBloq = hastaBloq, hastaBloq + 1, hastaBloq)) Then
                            bloqueado = True
                            Exit For
                        End If
                    End If
                End If
            Next j
        End If

        If bloqueado Then
            Me.ListBox1.AddItem apartamento & " - BLOQUEADO"
        Else
            Me.ListBox1.AddItem apartamento
        End If
    Next i

    Me.AceptarButton.SetFocus

    Exit Sub

SinDisponibles:
    MsgBox "No se pudo recuperar la lista de apartamentos disponibles.", vbExclamation
    Unload Me
End Sub

' === FUNCI�N AUXILIAR ===
Private Function ObtenerFilaDeApartamento(ByVal apartamento As String) As Long
    On Error Resume Next
    ObtenerFilaDeApartamento = 0
    
    Select Case Trim(apartamento)
        Case "Ap.2": ObtenerFilaDeApartamento = 2
        Case "Ap.4": ObtenerFilaDeApartamento = 4
        Case "Ap.5": ObtenerFilaDeApartamento = 6
        Case "Ap.6": ObtenerFilaDeApartamento = 8
        Case "Ap.7": ObtenerFilaDeApartamento = 10
        Case "Ap.8": ObtenerFilaDeApartamento = 12
        Case "Ap.9": ObtenerFilaDeApartamento = 14
        Case "Ap.10": ObtenerFilaDeApartamento = 16
        Case "Ap.12": ObtenerFilaDeApartamento = 18
    End Select
End Function
' Evento Click del bot�n Aceptar
Private Sub AceptarButton_Click()
    On Error GoTo ManejoErrores

    ' Variables para manejo de selecci�n
    Dim i As Long, seleccionadas As String
    Dim contadorSeleccionadas As Long  ' *** NUEVO: Contador de elementos seleccionados ***
    Dim seleccionArray() As String
    
    ' Variables para worksheets y datos
    Dim wsBloq As Worksheet, wsRes As Worksheet
    Dim fechaEntrada As Date, fechaSalida As Date
    Dim filaActual As Long
    
    ' Variables para verificaci�n de bloqueos
    Dim filaBloq As Long, arrBloq As Variant
    Dim j As Long, k As Long
    Dim haySolape As Boolean
    Dim mensajeSolape As String

    ' *** NUEVO: Construir string con habitaciones seleccionadas Y contar selecciones ***
    Dim itemStr As String
seleccionadas = ""
contadorSeleccionadas = 0
For i = 0 To Me.ListBox1.ListCount - 1
    If Me.ListBox1.Selected(i) Then
        contadorSeleccionadas = contadorSeleccionadas + 1
        itemStr = Me.ListBox1.List(i)
        If InStr(itemStr, " - ") > 0 Then
            itemStr = Split(itemStr, " - ")(0)
        End If
        If seleccionadas <> "" Then seleccionadas = seleccionadas & ", "
        seleccionadas = seleccionadas & itemStr
    End If
Next i

    ' *** NUEVO: Validar que solo se haya seleccionado UN apartamento ***
    If contadorSeleccionadas > 1 Then
        MsgBox "S�lo se puede seleccionar un apartamento por N� ORDEN. " & vbCrLf & _
               "Si el peticionario quiere m�s de un apartamento, se le debe adjudicar con otro N� ORDEN.", _
               vbExclamation, "Selecci�n m�ltiple no permitida"
        Exit Sub
    End If

    ' Validar que se haya seleccionado al menos una habitaci�n
    If seleccionadas = "" Then
        MsgBox "Debes seleccionar al menos un apartamento.", vbExclamation
        Exit Sub
    End If

    ' --- VALIDACI�N DE SOLAPES CON BLOQUEOS ---
    Set wsBloq = ThisWorkbook.Sheets("Apartamentos Bloqueados SOTO")
    Set wsRes = ThisWorkbook.Sheets("RESIDENCIA SOTO")
    
    ' Obtener fila actual y fechas desde la hoja de residencia
    filaActual = ActiveCell.Row
    
    ' Verificar que las fechas existan y sean v�lidas
    If Not IsDate(wsRes.Cells(filaActual, "L").Value) Or Not IsDate(wsRes.Cells(filaActual, "M").Value) Then
        MsgBox "Las fechas de entrada y salida deben estar definidas antes de seleccionar apartamentos.", vbExclamation
        Exit Sub
    End If
    
    fechaEntrada = wsRes.Cells(filaActual, "L").Value  ' Columna L: Fecha entrada
    fechaSalida = wsRes.Cells(filaActual, "M").Value   ' Columna M: Fecha salida

    ' Validar que la fecha de salida sea posterior a la de entrada
    If fechaSalida <= fechaEntrada Then
        MsgBox "La fecha de salida debe ser posterior a la fecha de entrada.", vbExclamation
        Exit Sub
    End If

    ' Obtener datos de bloqueos existentes
    filaBloq = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    If filaBloq >= 2 Then
        arrBloq = wsBloq.Range("A2:D" & filaBloq).Value  ' A:D = Apartamento, Desde, Hasta, Motivo
    Else
        arrBloq = Empty
    End If

    ' Verificar solapamientos por cada apartamento seleccionado
    seleccionArray = Split(seleccionadas, ",")
    haySolape = False
    mensajeSolape = ""

    If Not IsEmpty(arrBloq) Then
        For i = 0 To UBound(seleccionArray)
            For j = 1 To UBound(arrBloq, 1)
                ' Comparar nombres de apartamento (case-insensitive)
                If StrComp(Trim(seleccionArray(i)), Trim(arrBloq(j, 1)), vbTextCompare) = 0 Then
                    ' Validar que las fechas de bloqueo sean v�lidas
                    If IsDate(arrBloq(j, 2)) And IsDate(arrBloq(j, 3)) Then
                        Dim desdeBloq As Date, hastaBloq As Date
                        desdeBloq = arrBloq(j, 2)  ' Fecha inicio bloqueo
                        hastaBloq = arrBloq(j, 3)   ' Fecha fin bloqueo
                        
                        ' L�gica de detecci�n de solape:
                        ' NOT (La reserva termina ANTES del bloqueo O empieza DESPU�S del bloqueo)
                            If Not (fechaSalida <= desdeBloq Or fechaEntrada >= IIf(desdeBloq = hastaBloq, hastaBloq + 1, hastaBloq)) Then
                            haySolape = True
                            ' Construir mensaje detallado del conflicto
                            mensajeSolape = mensajeSolape & vbCrLf & _
                                "Apartamento: " & Trim(seleccionArray(i)) & _
                                " | Bloqueado del " & Format(desdeBloq, "dd/mm/yyyy") & _
                                " al " & Format(hastaBloq, "dd/mm/yyyy") & _
                                IIf(Not IsEmpty(arrBloq(j, 4)), " | Motivo: " & arrBloq(j, 4), "")
                        End If
                    End If
                End If
            Next j
        Next i
    End If

        ' Manejar casos con solape
    If haySolape Then
        Dim respuesta As VbMsgBoxResult
        respuesta = MsgBox("�Atenci�n! La actual selecci�n se solapa parcialmente con fechas bloqueadas:" & _
                           vbCrLf & mensajeSolape & vbCrLf & vbCrLf & _
                           "�Desea continuar con la selecci�n y eliminar el bloqueo?", _
                           vbExclamation + vbYesNo, "Solape con bloqueo")
                  
        If respuesta = vbNo Then
            Exit Sub  ' Abortar operaci�n
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
                                    wsBloq.Cells(j + 2, 1).Value = arrBloq(j, 1) ' Apartamento
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

    ' --- GUARDAR SELECCI�N ---
    ' Escribir la selecci�n en la celda correspondiente (Columna S para SOTO)
    wsRes.Cells(filaActual, "S").Value = seleccionadas
    
    ' --- RESOLUCION automatica: CONCEDIDA o SI segun dias de antelacion ---
    If IsDate(wsRes.Cells(filaActual, "B").Value) And IsDate(wsRes.Cells(filaActual, "L").Value) Then
        Dim fechaSolS As Date, umbralS As Long, diasAntS As Long
        fechaSolS = CDate(wsRes.Cells(filaActual, "B").Value)
        umbralS = 14 - Weekday(fechaSolS, vbMonday)
        diasAntS = DateDiff("d", fechaSolS, CDate(wsRes.Cells(filaActual, "L").Value))
        Application.enableEvents = False
        If diasAntS <= umbralS Then
            wsRes.Cells(filaActual, "P").Value = "CONCEDIDA"
        Else
            wsRes.Cells(filaActual, "P").Value = "SI"
        End If
        Application.enableEvents = True
    End If
    
    ' Cerrar el formulario
    Unload Me
    Exit Sub

ManejoErrores:
    MsgBox "Se produjo un error: " & Err.Description, vbCritical
    Exit Sub
End Sub

' Evento Click del boton Cancelar
Private Sub CancelarButton_Click()
    ' --- RESOLUCION automatica por cancelacion: DENEGADA o NO segun dias de antelacion ---
    Dim wsResCS As Worksheet
    Dim filaCS As Long
    Set wsResCS = ThisWorkbook.Sheets("RESIDENCIA SOTO")
    filaCS = ActiveCell.Row
    Dim valorActualCS As String
    valorActualCS = UCase(Trim(wsResCS.Cells(filaCS, "P").Value))
    If valorActualCS <> "SI" And valorActualCS <> "CONCEDIDA" And _
       valorActualCS <> "REEVALUADA" And valorActualCS <> "RENUNCIA" Then
        If IsDate(wsResCS.Cells(filaCS, "B").Value) And IsDate(wsResCS.Cells(filaCS, "L").Value) Then
            Dim fechaSolCS As Date, umbralCS As Long, diasAntCS As Long
            fechaSolCS = CDate(wsResCS.Cells(filaCS, "B").Value)
            umbralCS = 14 - Weekday(fechaSolCS, vbMonday)
            diasAntCS = DateDiff("d", fechaSolCS, CDate(wsResCS.Cells(filaCS, "L").Value))
            Application.enableEvents = False
            If diasAntCS <= umbralCS Then
                wsResCS.Cells(filaCS, "P").Value = "DENEGADA"
            Else
                wsResCS.Cells(filaCS, "P").Value = "NO"
            End If
            Application.enableEvents = True
        End If
    End If
    Unload Me  ' Cerrar el formulario sin acciones
End Sub

