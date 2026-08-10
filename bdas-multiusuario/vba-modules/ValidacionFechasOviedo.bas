Attribute VB_Name = "ValidacionFechasOviedo"
'=================================================================================
' M�dulo: ValidacionFechasOviedo
' Descripci�n:
'   Determina qu� habitaciones est�n disponibles en la hoja "RESIDENCIA OVIEDO"
'   para un rango de fechas y una separaci�n m�nima entre reservas.
'   Utiliza arrays en memoria para m�xima eficiencia y muestra un formulario
'   de selecci�n si hay habitaciones libres. Define el nombre "HabitacionesDisponiblesOVIEDO"
'   con la lista de habitaciones libres, �til para validaciones o formularios.
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' VerificarDisponibilidadOviedo
'   - fechaEntrada, fechaSalida: rango de fechas a comprobar (tipo Date)
'   - filaActual: (no usado en la l�gica actual, reservado para futuras ampliaciones)
'   - Separacion: d�as de separaci�n requeridos entre reservas (no puede ser negativo)
'---------------------------------------------------------------------------------
Public Sub VerificarDisponibilidadOviedo(fechaEntrada As Date, fechaSalida As Date, filaActual As Long, Separacion As Long)
    On Error GoTo ManejoErrores

    '---------------------- Validaci�n de par�metro Separacion -----------------------
    If Separacion < 0 Then
        MsgBox "Error: El valor de separaci�n no puede ser negativo.", vbCritical
        Exit Sub
    End If

    '---------------------- Inicializaci�n de variables y arrays ---------------------
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("RESIDENCIA OVIEDO")
    
    ' --- VERIFICAR DUPLICADOS EN LOS �LTIMOS 30 D͍AS ---
    Call VerificarDuplicadoUltimos30Dias(ws, filaActual)
    
    ' Si la solicitud est� DESESTIMADA o vac�a tras la validaci�n, salimos tempranamente
    If UCase(Trim(ws.Cells(filaActual, "P").Value)) = "DESESTIMADA" Then Exit Sub
    If IsEmpty(ws.Cells(filaActual, "L").Value) Or IsEmpty(ws.Cells(filaActual, "M").Value) Then Exit Sub
    Dim wsCal As Worksheet
    On Error Resume Next
    Set wsCal = ThisWorkbook.Sheets("Calendario OVIEDO")
    On Error GoTo ManejoErrores

    Dim habitacionesDisponibles2 As String
    habitacionesDisponibles2 = ""

    ' Array de habitaciones a comprobar (puede personalizarse f�cilmente)
    Dim habitaciones() As String
    habitaciones = Split("1,2,3,4,5,6,7,8,9,10,11,12,13,14,15", ",")

    '---------------------- Determinaci�n de la �ltima fila --------------------------
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    If ultimaFila < 2 Then
        MsgBox "No hay datos en la hoja RESIDENCIA OVIEDO.", vbInformation
        Exit Sub
    End If

    '---------------------- Carga de datos en arrays para m�xima velocidad -----------
    Dim arrEntradas As Variant, arrSalidas As Variant, arrHabitaciones As Variant
    
    ' MANEJO ROBUSTO DE DATOS PARA CASOS DE 1 FILA
    If ultimaFila = 2 Then
        ReDim arrEntradas(1 To 1, 1 To 1)
        arrEntradas(1, 1) = ws.Range("L2").Value
        
        ReDim arrSalidas(1 To 1, 1 To 1)
        arrSalidas(1, 1) = ws.Range("M2").Value
        
        ReDim arrHabitaciones(1 To 1, 1 To 1)
        arrHabitaciones(1, 1) = ws.Range("T2").Value
    Else
        arrEntradas = ws.Range("L2:L" & ultimaFila).Value        ' Fechas de entrada de reservas
        arrSalidas = ws.Range("M2:M" & ultimaFila).Value         ' Fechas de salida de reservas
        arrHabitaciones = ws.Range("T2:T" & ultimaFila).Value    ' Habitaciones adjudicadas por reserva
    End If

    '---------------------- B�squeda de habitaciones disponibles ---------------------
    Dim k As Long, i As Long, j As Long
    Dim habitacion As String
    Dim habitacionLibre As Boolean
    Dim habitacionesAdjudicadas As Variant
    Dim fechaEntradaExistente As Date, fechaSalidaExistente As Date

    For k = LBound(habitaciones) To UBound(habitaciones)
        habitacion = Trim(habitaciones(k))
        habitacionLibre = True   ' Suponemos disponible al principio

        ' Recorremos todas las reservas existentes
        For i = 1 To UBound(arrEntradas, 1)
            ' Excluir la fila actual que se est� editando
            If (ultimaFila = 2 And filaActual = 2) Or (i + 1 = filaActual) Then GoTo SiguienteReserva

            ' Si alguna fecha no es v�lida, saltamos ese registro (protege ante errores de datos)
            If Not (IsDate(arrEntradas(i, 1)) And IsDate(arrSalidas(i, 1))) Then GoTo SiguienteReserva

            fechaEntradaExistente = arrEntradas(i, 1)
            fechaSalidaExistente = arrSalidas(i, 1)

            ' Leemos habitaciones adjudicadas en la reserva (pueden ser varias, separadas por coma)
            If Trim(arrHabitaciones(i, 1) & "") <> "" Then
                habitacionesAdjudicadas = Split(Replace(arrHabitaciones(i, 1), " ", ""), ",")
            Else
                habitacionesAdjudicadas = Array()
            End If

            ' Comprobamos si la habitaci�n est� adjudicada en la reserva actual
            For j = LBound(habitacionesAdjudicadas) To UBound(habitacionesAdjudicadas)
                If Trim(habitacionesAdjudicadas(j)) = habitacion Then
                    ' Comprobaci�n de solape entre reservas, considerando la separaci�n m�nima
                    ' Si las fechas NO est�n completamente separadas por al menos 'Separacion' d�as,
                    ' la habitaci�n NO est� disponible
                    If Not (fechaSalida <= DateAdd("d", -Separacion, fechaEntradaExistente) Or _
                            fechaEntrada >= DateAdd("d", Separacion, fechaSalidaExistente)) Then
                        habitacionLibre = False
                        Exit For
                    End If
                End If
            Next j

            If Not habitacionLibre Then Exit For
SiguienteReserva:
        Next i

        ' Si no hay conflictos, comprobar que no est� oculta en Calendario
        If habitacionLibre Then
            Dim filaCalHab As Long
            filaCalHab = 0
            If Not wsCal Is Nothing Then
                filaCalHab = ObtenerFilaDeHabitacionOviedo(habitacion)
            End If
            If filaCalHab > 0 Then
                If Not wsCal.Rows(filaCalHab).Hidden Then
                    habitacionesDisponibles2 = habitacionesDisponibles2 & habitacion & ", "
                End If
            Else
                habitacionesDisponibles2 = habitacionesDisponibles2 & habitacion & ", "
            End If
        End If
    Next k

    '---------------------- Resultados finales: visualizaci�n y nombre definido -------
    If Len(habitacionesDisponibles2) > 0 Then
        ' Quita la �ltima coma y espacio
        habitacionesDisponibles2 = Left(habitacionesDisponibles2, Len(habitacionesDisponibles2) - 2)
        
        Dim arrDisponibles() As String
        arrDisponibles = Split(habitacionesDisponibles2, ", ")
        
                Dim paxVal As Long
        paxVal = val(ws.Cells(filaActual, "O").Value)
        
        Dim idealList As Collection
        Dim greaterList As Collection
        Dim insufList As Collection
        Dim capTotal As Long
        Dim hab As Variant
        Dim capHab As Long
        
        Set idealList = New Collection
        Set greaterList = New Collection
        Set insufList = New Collection
        capTotal = 0
        
        For Each hab In arrDisponibles
            capHab = ObtenerCapacidadOviedo(CStr(hab))
            capTotal = capTotal + capHab
            If capHab >= paxVal Then
                ' Aplicar matizaciones de capacidad ideal (Excel 2016 en Espa�ol)
                If paxVal = 1 And capHab > 2 Then
                    greaterList.Add CStr(hab)
                ElseIf paxVal = 2 And capHab > 3 Then
                    greaterList.Add CStr(hab)
                Else
                    idealList.Add CStr(hab)
                End If
            Else
                insufList.Add CStr(hab)
            End If
        Next hab
        
        Dim finalShowList As String
        finalShowList = ""
        
        If idealList.Count > 0 Then
            ' Caso A: Hay habitaciones con la capacidad ideal
            Dim idxIdeal As Long
            For idxIdeal = 1 To idealList.Count
                finalShowList = finalShowList & idealList(idxIdeal) & ", "
            Next idxIdeal
            
            ' Regla Especial Oviedo: Si PAX = 4 y hay al menos 2 dobles libres, se a�aden como alternativa
            If paxVal = 4 Then
                Dim cantDobles As Long
                cantDobles = 0
                Dim habInsuf As Variant
                For Each habInsuf In insufList
                    If ObtenerCapacidadOviedo(CStr(habInsuf)) = 2 Then
                        cantDobles = cantDobles + 1
                    End If
                Next habInsuf
                
                If cantDobles >= 2 Then
                    For Each habInsuf In insufList
                        If ObtenerCapacidadOviedo(CStr(habInsuf)) = 2 Then
                            finalShowList = finalShowList & CStr(habInsuf) & ", "
                        End If
                    Next habInsuf
                End If
            End If
            
            finalShowList = Left(finalShowList, Len(finalShowList) - 2)
            
            On Error Resume Next
            ThisWorkbook.Names("HabitacionesDisponiblesOVIEDO").Delete
            On Error GoTo 0
            
            ThisWorkbook.Names.Add Name:="HabitacionesDisponiblesOVIEDO", RefersTo:="=" & Chr(34) & finalShowList & Chr(34)
            
            Application.enableEvents = True
            Application.screenUpdating = True
            DoEvents
            SeleccionHabitacionOVIEDO.Show
            
        ElseIf greaterList.Count > 0 Then
            ' Caso A (Matizado): No hay ideales pero s� de mayor capacidad (para PAX = 1 o 2)
            Dim pregMayor As VbMsgBoxResult
            pregMayor = MsgBox("No existen habitaciones de esa capacidad, �desea adjudicar una habitaci�n con mayor capacidad?", vbQuestion + vbYesNo, "Habitaci�n de mayor capacidad")
            If pregMayor = vbYes Then
                Dim idxGreater As Long
                For idxGreater = 1 To greaterList.Count
                    finalShowList = finalShowList & greaterList(idxGreater) & ", "
                Next idxGreater
                finalShowList = Left(finalShowList, Len(finalShowList) - 2)
                
                On Error Resume Next
                ThisWorkbook.Names("HabitacionesDisponiblesOVIEDO").Delete
                On Error GoTo 0
                
                ThisWorkbook.Names.Add Name:="HabitacionesDisponiblesOVIEDO", RefersTo:="=" & Chr(34) & finalShowList & Chr(34)
                
                Application.enableEvents = True
                Application.screenUpdating = True
                DoEvents
                SeleccionHabitacionOVIEDO.Show
            Else
                ' Si selecciona NO, se marca como NO o DENEGADA seg�n la antelaci�n de la solicitud
                Call EscribirResolucionDenegadaOviedo(ws, filaActual, fechaEntrada)
                Exit Sub
            End If
            
        Else
            ' Paso 2: Asignaci�n M�ltiple (Combinaciones)
            If capTotal >= paxVal Then
                Dim pregunta As VbMsgBoxResult
                pregunta = MsgBox("No existe disponible ninguna habitaci�n de esa capacidad para alojar a los " & paxVal & " PAX." & vbCrLf & vbCrLf & _
                                  "�Desea asignar m�s de una habitaci�n para esta reserva?", vbQuestion + vbYesNo, "Habitaci�n insuficiente")
                If pregunta = vbYes Then
                    On Error Resume Next
                    ThisWorkbook.Names("HabitacionesDisponiblesOVIEDO").Delete
                    On Error GoTo 0
                    
                    ThisWorkbook.Names.Add Name:="HabitacionesDisponiblesOVIEDO", RefersTo:="=" & Chr(34) & habitacionesDisponibles2 & Chr(34)
                    
                    Application.enableEvents = True
                    Application.screenUpdating = True
                    DoEvents
                    SeleccionHabitacionOVIEDO.Show
                Else
                    ' Usuario rechaza asignacion multiple -> DENEGADA o NO segun antelacion
                    Call EscribirResolucionDenegadaOviedo(ws, filaActual, fechaEntrada)
                    Exit Sub
                End If
            Else
                ' Paso 3: Capacidad Insuficiente Absoluta
                MsgBox "No existe disponibilidad de alojamiento para la solicitud grabada.", vbExclamation, "Capacidad Insuficiente"
                Call EscribirResolucionDenegadaOviedo(ws, filaActual, fechaEntrada)
                Exit Sub
            End If
        End If
    Else
        ' Paso 3: Capacidad Insuficiente Absoluta (sin habitaciones disponibles)
        MsgBox "No existe disponibilidad de alojamiento para la solicitud grabada.", vbExclamation, "Capacidad Insuficiente"
        Call EscribirResolucionDenegadaOviedo(ws, filaActual, fechaEntrada)
    End If

    Exit Sub

ManejoErrores:
    ' Captura cualquier error inesperado y lo muestra al usuario
    MsgBox "Se produjo un error inesperado: " & Err.Description, vbCritical
End Sub

Public Function ObtenerCapacidadOviedo(ByVal hab As String) As Long
    Select Case UCase(Trim(hab))
        Case "1", "2", "3": ObtenerCapacidadOviedo = 3
        Case "4", "5", "7", "8", "9", "10", "11", "13", "14", "15": ObtenerCapacidadOviedo = 2
        Case "6": ObtenerCapacidadOviedo = 1
        Case "12": ObtenerCapacidadOviedo = 4
        Case Else: ObtenerCapacidadOviedo = 2
    End Select
End Function

Public Function ObtenerFilaDeHabitacionOviedo(ByVal habitacion As String) As Long
    ObtenerFilaDeHabitacionOviedo = 0
    Select Case Trim(habitacion)
        Case "1":  ObtenerFilaDeHabitacionOviedo = 2
        Case "2":  ObtenerFilaDeHabitacionOviedo = 4
        Case "3":  ObtenerFilaDeHabitacionOviedo = 6
        Case "4":  ObtenerFilaDeHabitacionOviedo = 8
        Case "5":  ObtenerFilaDeHabitacionOviedo = 10
        Case "6":  ObtenerFilaDeHabitacionOviedo = 12
        Case "7":  ObtenerFilaDeHabitacionOviedo = 14
        Case "8":  ObtenerFilaDeHabitacionOviedo = 16
        Case "9":  ObtenerFilaDeHabitacionOviedo = 18
        Case "10": ObtenerFilaDeHabitacionOviedo = 20
        Case "11": ObtenerFilaDeHabitacionOviedo = 22
        Case "12": ObtenerFilaDeHabitacionOviedo = 24
        Case "13": ObtenerFilaDeHabitacionOviedo = 26
        Case "14": ObtenerFilaDeHabitacionOviedo = 28
        Case "15": ObtenerFilaDeHabitacionOviedo = 30
    End Select
End Function

'------------------------------------------------------------------------------------------
' EscribirResolucionDenegadaOviedo
'   Escribe DENEGADA o NO en col. P segun los dias de antelacion entre solicitud y entrada.
'   El umbral varia segun el dia de la semana de la fecha de solicitud (col. B):
'     Lunes=13, Martes=12, Mier=11, Jue=10, Vie=9, Sab=8, Dom=7
'   Si col. P ya contiene un valor definitivo (SI/CONCEDIDA/REEVALUADA/RENUNCIA), no sobreescribe.
'------------------------------------------------------------------------------------------
Public Sub EscribirResolucionDenegadaOviedo(ws As Worksheet, fila As Long, fechaEntrada As Date)
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
