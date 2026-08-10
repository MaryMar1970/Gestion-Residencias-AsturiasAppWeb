Attribute VB_Name = "ValidacionFechasGijon"
Option Explicit
' Modulo ValidacionFechasGijon v14.11.3
' =========================================================================================
' FUNCIONES AUXILIARES DE SOLAPE
' =========================================================================================

'------------------------------------------------------------------------------------------
' HaySolapeReservas:
'   Determina si dos reservas se solapan en el tiempo.
'   Devuelve True si hay solape, False si no lo hay.
'   - inicio1, fin1: Fechas de inicio y fin del primer periodo
'   - inicio2, fin2: Fechas de inicio y fin del segundo periodo
'------------------------------------------------------------------------------------------
Private Function HaySolapeReservas(inicio1 As Date, fin1 As Date, inicio2 As Date, fin2 As Date) As Boolean
    ' Si un periodo termina antes de que el otro empiece, no hay solape
    ' Si un periodo empieza despu�s de que el otro termine, no hay solape
    ' En cualquier otro caso, hay solape
    HaySolapeReservas = Not (fin1 <= inicio2 Or inicio1 >= fin2)
End Function

'------------------------------------------------------------------------------------------
' HaySolapeBloqueo:
'   Determina si una reserva solapa un bloqueo de habitaci�n.
'   Devuelve True si hay solape, False si no lo hay.
'   - reservaInicio, reservaFin: Fechas de la reserva a comprobar
'   - bloqueoInicio, bloqueoFin: Fechas del bloqueo
'   * Suma 1 d�a a bloqueoFin para incluir el d�a de salida del bloqueo
'------------------------------------------------------------------------------------------
Public Function HaySolapeBloqueo(reservaInicio As Date, reservaFin As Date, bloqueoInicio As Date, bloqueoFin As Date) As Boolean
    Dim limiteFin As Date
    If bloqueoInicio = bloqueoFin Then
        limiteFin = bloqueoFin + 1
    Else
        limiteFin = bloqueoFin
    End If
    HaySolapeBloqueo = Not (reservaFin <= bloqueoInicio Or reservaInicio >= limiteFin)
End Function

' =========================================================================================
' VERIFICAR DISPONIBILIDAD DE HABITACIONES EN GIJ�N
' =========================================================================================

'------------------------------------------------------------------------------------------
' VerificarDisponibilidadGijon:
'   Determina qu� habitaciones est�n libres para un rango de fechas en GIJ�N,
'   considerando reservas existentes y bloqueos.
'   - fechaEntrada, fechaSalida: Rango de fechas a comprobar (tipo Date)
'   - filaActual: Fila activa (para excluirla de la validaci�n)
'   - Separacion: Separaci�n entre reservas (en d�as) para evitar solapes
'
'   Muestra un formulario de selecci�n si hay habitaciones disponibles;
'   si no, avisa al usuario.
'   Crea el nombre definido "HabitacionesDisponiblesGIJON" con la lista de habitaciones libres.
'------------------------------------------------------------------------------------------
Public Sub VerificarDisponibilidadGijon(fechaEntrada As Date, fechaSalida As Date, filaActual As Long, Separacion As Long)
    On Error GoTo ManejoErrores

    '--------------------------- Variables -------------------------------
    Dim ws As Worksheet                    ' Hoja de reservas (RESIDENCIA GIJ�N)
    Dim wsBloq As Worksheet                ' Hoja de bloqueos (Habitaciones Bloqueadas GIJ�N)
    Dim habitaciones() As String           ' Array de nombres de habitaciones a comprobar
    Dim habitacionLibre As Boolean         ' Indica si la habitaci�n est� libre
    Dim bloqueada As Boolean               ' Indica si la habitaci�n est� bloqueada
    Dim habitacion As String               ' Nombre de habitaci�n en el bucle
    Dim i As Long, j As Long, k As Long    ' ͍ndices para bucles
    Dim fechaEntradaExistente As Date, fechaSalidaExistente As Date
    Dim habitacionesAdjudicadas As Variant ' Habitaciones adjudicadas por reserva
    Dim ultimaFila As Long, filaBloq As Long
    Dim arrEntrada As Variant, arrSalida As Variant, arrHabitaciones As Variant
    Dim arrBloq As Variant
    Dim bloqueos As Collection             ' Colecci�n de bloqueos en memoria
    Dim bloqArr As Variant                 ' Array temporal para un bloqueo
    Dim habitacionesDisponibles As Collection   ' Colecci�n de habitaciones finalmente libres

    '--------------------------- Inicializaciones ------------------------
    Set ws = ThisWorkbook.Sheets("RESIDENCIA GIJ�N")
    
    ' --- VERIFICAR DUPLICADOS EN LOS �LTIMOS 30 D͍AS ---
    Call VerificarDuplicadoUltimos30Dias(ws, filaActual)
    
    ' Si la solicitud est� DESESTIMADA o vac�a tras la validaci�n, salimos tempranamente
    If UCase(Trim(ws.Cells(filaActual, "P").Value)) = "DESESTIMADA" Then Exit Sub
    If IsEmpty(ws.Cells(filaActual, "L").Value) Or IsEmpty(ws.Cells(filaActual, "M").Value) Then Exit Sub
    Dim wsCal As Worksheet
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets("Calendario GIJ�N")
    On Error GoTo ManejoErrores
    
    On Error Resume Next
    Set wsBloq = ThisWorkbook.Sheets("Habitaciones Bloqueadas GIJ�N")
    On Error GoTo ManejoErrores
    
    habitaciones = Split("1,2,3,4,5,6,7,Of.1,Of.2,Of.3,Est.1,Est.2,Est.3", ",") ' Habitaciones a comprobar
    Set habitacionesDisponibles = New Collection

    '--------------------------- Cargar bloqueos en memoria ------------------------
    Set bloqueos = New Collection
    
    If Not wsBloq Is Nothing Then
        On Error Resume Next
        filaBloq = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
        If Err.Number = 0 And filaBloq >= 2 Then
            arrBloq = wsBloq.Range("A2:C" & filaBloq).Value   ' Carga la tabla de bloqueos (A: habitaci�n, B: inicio, C: fin)
            If IsArray(arrBloq) Then
                For i = 1 To UBound(arrBloq, 1)
                    ' Validaci�n: solo a�ade bloqueos si hay habitaci�n y fechas v�lidas
                    If arrBloq(i, 1) <> "" And IsDate(arrBloq(i, 2)) And IsDate(arrBloq(i, 3)) Then
                        bloqueos.Add Array(CStr(arrBloq(i, 1)), CDate(arrBloq(i, 2)), CDate(arrBloq(i, 3)))
                    End If
                Next i
            End If
        End If
        On Error GoTo ManejoErrores
    End If

    '--------------------------- Cargar reservas actuales ------------------------
    ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    If ultimaFila < 2 Then
        ' Si no hay reservas, todas las habitaciones est�n disponibles (salvo bloqueos)
        ' Continuamos sin reservas
    Else
        ' MANEJO ROBUSTO DE DATOS PARA CASOS DE 1 FILA
        If ultimaFila = 2 Then
            ' Caso especial: solo 1 fila de datos (fila 2)
            ReDim arrEntrada(1 To 1, 1 To 1)
            arrEntrada(1, 1) = ws.Range("L2").Value
            
            ReDim arrSalida(1 To 1, 1 To 1)
            arrSalida(1, 1) = ws.Range("M2").Value
            
            ReDim arrHabitaciones(1 To 1, 1 To 1)
            arrHabitaciones(1, 1) = ws.Range("S2").Value
        Else
            ' Caso normal: m�ltiples filas de datos
            arrEntrada = ws.Range("L2:L" & ultimaFila).Value
            arrSalida = ws.Range("M2:M" & ultimaFila).Value
            arrHabitaciones = ws.Range("S2:S" & ultimaFila).Value
        End If
    End If

    '--------------------------- B�squeda de habitaciones disponibles ------------------------
    For k = LBound(habitaciones) To UBound(habitaciones)
        habitacion = Trim(habitaciones(k))
        habitacionLibre = True   ' Suponemos que est� libre al principio
        bloqueada = False

        ' 1. Revisar bloqueos
        For j = 1 To bloqueos.Count
            bloqArr = bloqueos(j)
            ' Si la habitaci�n coincide y hay solape con el bloqueo, se marca como bloqueada
            If StrComp(bloqArr(0), habitacion, vbTextCompare) = 0 Then
                If HaySolapeBloqueo(fechaEntrada, fechaSalida, CDate(bloqArr(1)), CDate(bloqArr(2))) Then
                    bloqueada = True
                    Exit For
                End If
            End If
        Next j
'        If bloqueada Then GoTo SiguienteHabitacion  ' Si est� bloqueada, pasa a la siguiente habitaci�n

        ' 2. Revisar solapamientos con reservas existentes (solo si hay reservas)
        If ultimaFila >= 2 Then
            For i = 1 To UBound(arrEntrada, 1)
                ' Excluir la fila actual que se est� editando
                If (ultimaFila = 2 And filaActual = 2) Or (i + 1 = filaActual) Then GoTo SiguienteReserva
                
                ' Validar que existen fechas en la reserva
                If Not IsDate(arrEntrada(i, 1)) Or Not IsDate(arrSalida(i, 1)) Then GoTo SiguienteReserva
                
                fechaEntradaExistente = arrEntrada(i, 1)
                fechaSalidaExistente = arrSalida(i, 1)
                
                If arrHabitaciones(i, 1) <> "" Then
                    ' Divide las habitaciones adjudicadas (pueden ser varias separadas por coma)
                    habitacionesAdjudicadas = Split(Replace(arrHabitaciones(i, 1), " ", ""), ",")
                    For j = LBound(habitacionesAdjudicadas) To UBound(habitacionesAdjudicadas)
                        ' Si la habitaci�n coincide y hay solape de fechas, no est� libre
                        If Trim(habitacionesAdjudicadas(j)) = habitacion Then
                            ' Comprobaci�n de solape considerando separaci�n m�nima
                            If HaySolapeReservas(fechaEntrada, fechaSalida, _
                                DateAdd("d", -Separacion, fechaEntradaExistente), _
                                DateAdd("d", Separacion, fechaSalidaExistente)) Then
                                habitacionLibre = False
                                Exit For
                            End If
                        End If
                    Next j
                    If Not habitacionLibre Then Exit For
                End If
SiguienteReserva:
            Next i
        End If

        ' Si la habitaci�n sigue libre, comprobar que no est� oculta en Calendario
        If habitacionLibre Then
            Dim filaCalHab As Long
            filaCalHab = 0
            If Not wsCal Is Nothing Then
                filaCalHab = ObtenerFilaDeHabitacionGijon(habitacion)
            End If
            If filaCalHab > 0 Then
                If Not wsCal.Rows(filaCalHab).Hidden Then
                    habitacionesDisponibles.Add habitacion
                End If
            Else
                habitacionesDisponibles.Add habitacion
            End If
        End If

SiguienteHabitacion:
        ' Siguiente habitaci�n del bucle principal
    Next k

    '--------------------------- Resultado: mostrar selecci�n o mensaje ------------------------
    If habitacionesDisponibles.Count > 0 Then
        ' Generar la lista de habitaciones libres como texto separado por comas
        Dim listaDisponibles As String
        listaDisponibles = ""
        For i = 1 To habitacionesDisponibles.Count
            listaDisponibles = listaDisponibles & habitacionesDisponibles(i) & ", "
        Next i
        listaDisponibles = Left(listaDisponibles, Len(listaDisponibles) - 2)

        Dim arrDisponibles() As String
        arrDisponibles = Split(listaDisponibles, ", ")
        
        Dim paxVal As Long
        paxVal = val(ws.Cells(filaActual, "O").Value)
        
        Dim sufList As Collection
        Dim insufList As Collection
        Dim capTotal As Long
        Dim hab As Variant
        Dim capHab As Long
        
        Set sufList = New Collection
        Set insufList = New Collection
        capTotal = 0
        
        For Each hab In arrDisponibles
            capHab = ObtenerCapacidadGijon(CStr(hab))
            capTotal = capTotal + capHab
            If capHab >= paxVal Then
                sufList.Add CStr(hab)
            Else
                insufList.Add CStr(hab)
            End If
        Next hab
        
        Dim finalShowList As String
        finalShowList = ""
        
        If sufList.Count > 0 Then
            ' Paso 1: Habitaciones Suficientes �nicas
            Dim idxSuf As Long
            For idxSuf = 1 To sufList.Count
                finalShowList = finalShowList & sufList(idxSuf) & ", "
            Next idxSuf
            finalShowList = Left(finalShowList, Len(finalShowList) - 2)
            
            On Error Resume Next
            ThisWorkbook.Names("HabitacionesDisponiblesGIJON").Delete
            On Error GoTo 0
            
            ThisWorkbook.Names.Add Name:="HabitacionesDisponiblesGIJON", RefersTo:="=" & Chr(34) & finalShowList & Chr(34)
            
            Application.enableEvents = True
            Application.screenUpdating = True
            DoEvents
            SeleccionHabitacionGIJON.Show
            
        Else
            ' Paso 2: Asignaci�n M�ltiple (Combinaciones)
            If capTotal >= paxVal Then
                Dim pregunta As VbMsgBoxResult
                pregunta = MsgBox("No existe disponible ninguna habitaci�n de esa capacidad para alojar a los " & paxVal & " PAX." & vbCrLf & vbCrLf & _
                                  "�Desea asignar m�s de una habitaci�n para esta reserva?", vbQuestion + vbYesNo, "Habitaci�n insuficiente")
                If pregunta = vbYes Then
                    On Error Resume Next
                    ThisWorkbook.Names("HabitacionesDisponiblesGIJON").Delete
                    On Error GoTo 0
                    
                    ThisWorkbook.Names.Add Name:="HabitacionesDisponiblesGIJON", RefersTo:="=" & Chr(34) & listaDisponibles & Chr(34)
                    
                    Application.enableEvents = True
                    Application.screenUpdating = True
                    DoEvents
                    SeleccionHabitacionGIJON.Show
                Else
                    ' Usuario rechaza asignacion multiple -> DENEGADA o NO segun antelacion
                    Call EscribirResolucionDenegada(ws, filaActual, fechaEntrada)
                    Exit Sub
                End If
            Else
                ' Paso 3: Capacidad Insuficiente Absoluta
                MsgBox "No existe disponibilidad de alojamiento para la solicitud grabada.", vbExclamation, "Capacidad Insuficiente"
                Call EscribirResolucionDenegada(ws, filaActual, fechaEntrada)
                Exit Sub
            End If
        End If
    Else
        ' Paso 3: Capacidad Insuficiente Absoluta (sin habitaciones disponibles)
        MsgBox "No existe disponibilidad de alojamiento para la solicitud grabada.", vbExclamation, "Capacidad Insuficiente"
        Call EscribirResolucionDenegada(ws, filaActual, fechaEntrada)
    End If

    Exit Sub

ManejoErrores:
    ' Captura cualquier error inesperado y lo muestra al usuario
    MsgBox "Se produjo un error en VerificarDisponibilidadGijon: " & Err.Description & _
        vbCrLf & "Origen: " & Err.Source, vbCritical
End Sub

Public Function ObtenerCapacidadGijon(ByVal hab As String) As Long
    Select Case UCase(Trim(hab))
        Case "4", "EST.1", "EST.2", "EST.3": ObtenerCapacidadGijon = 1
        Case "1", "2", "3", "5", "6", "7", "OF.1", "OF.2", "OF.3": ObtenerCapacidadGijon = 2
        Case Else: ObtenerCapacidadGijon = 0
    End Select
End Function

Public Function ObtenerFilaDeHabitacionGijon(ByVal habitacion As String) As Long
    ObtenerFilaDeHabitacionGijon = 0
    Select Case Trim(habitacion)
        Case "1":     ObtenerFilaDeHabitacionGijon = 2
        Case "2":     ObtenerFilaDeHabitacionGijon = 4
        Case "3":     ObtenerFilaDeHabitacionGijon = 6
        Case "4":     ObtenerFilaDeHabitacionGijon = 8
        Case "5":     ObtenerFilaDeHabitacionGijon = 10
        Case "6":     ObtenerFilaDeHabitacionGijon = 12
        Case "7":     ObtenerFilaDeHabitacionGijon = 14
        Case "Of.1":  ObtenerFilaDeHabitacionGijon = 16
        Case "Of.2":  ObtenerFilaDeHabitacionGijon = 18
        Case "Of.3":  ObtenerFilaDeHabitacionGijon = 20
        Case "Est.1": ObtenerFilaDeHabitacionGijon = 22
        Case "Est.2": ObtenerFilaDeHabitacionGijon = 24
        Case "Est.3": ObtenerFilaDeHabitacionGijon = 26
    End Select
End Function

'------------------------------------------------------------------------------------------
' EscribirResolucionDenegada (GIJON)
'   Escribe DENEGADA o NO en col. P segun los dias de antelacion entre solicitud y entrada.
'   El umbral varia segun el dia de la semana de la fecha de solicitud (col. B):
'     Lunes=13, Martes=12, Mier=11, Jue=10, Vie=9, Sab=8, Dom=7
'   Si col. P ya contiene un valor definitivo (SI/CONCEDIDA/REEVALUADA/RENUNCIA), no sobreescribe.
'------------------------------------------------------------------------------------------
Public Sub EscribirResolucionDenegada(ws As Worksheet, fila As Long, fechaEntrada As Date)
    Dim valorActual As String
    valorActual = UCase(Trim(ws.Cells(fila, "P").Value))
' Guard clause: preguntar antes de sobreescribir resoluciones definitivas
    If valorActual = "SI" Or valorActual = "CONCEDIDA" Or _
       valorActual = "REEVALUADA" Or valorActual = "RENUNCIA" Then
        Dim resp As VbMsgBoxResult
        resp = MsgBox("La celda ya contiene '" & valorActual & "'." & vbCrLf & _
                      "�Desea sobreescribir la resoluci�n?", _
                      vbYesNo + vbQuestion, "Confirmar sobreescritura")
        If resp = vbNo Then Exit Sub
    End If
    
    If Not IsDate(ws.Cells(fila, "B").Value) Then Exit Sub
    
    Dim fechaSolicitud As Date
    fechaSolicitud = CDate(ws.Cells(fila, "B").Value)
    
    ' Umbral variable segun dia de la semana (L=13, M=12, X=11, J=10, V=9, S=8, D=7)
    Dim umbral As Long
    umbral = 14 - Weekday(fechaSolicitud, vbMonday)
    
    Dim diasAnt As Long
    diasAnt = DateDiff("d", fechaSolicitud, fechaEntrada)
    
        On Error GoTo SafeExit
    Application.enableEvents = False
    If diasAnt <= umbral Then
        ws.Cells(fila, "P").Value = "DENEGADA"
    Else
        ws.Cells(fila, "P").Value = "NO"
    End If
SafeExit:
    Application.enableEvents = True
End Sub

