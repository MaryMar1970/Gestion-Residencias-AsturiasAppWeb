Attribute VB_Name = "modCapacidadGijon"
Option Explicit

Public Function ReevaluacionViableGijonPorCapacidad( _
        ByVal arrHabLib As Variant, _
        ByVal paxSolicitados As Long) As Boolean
    
    Dim paxRestante As Long
    Dim capacidadHab As Long
    Dim i As Long
    Dim doblesUsadasComoIndiv As Long
    
    Debug.Print "=== ReevaluacionViableGijonPorCapacidad ==="
    Debug.Print "PAX Solicitados: " & paxSolicitados
    
    paxRestante = paxSolicitados
    doblesUsadasComoIndiv = 0
    
    For i = LBound(arrHabLib) To UBound(arrHabLib)
        Debug.Print "Habitaci�n[" & i & "]: " & arrHabLib(i)
        
        Select Case UCase(arrHabLib(i))
            Case "IND"
                capacidadHab = 1
            Case "DOB"
                If paxRestante = 1 Then
                    capacidadHab = 1
                    doblesUsadasComoIndiv = doblesUsadasComoIndiv + 1
                Else
                    capacidadHab = 2
                End If
            Case Else
                capacidadHab = 0
        End Select
        
        Debug.Print "  Capacidad asignada: " & capacidadHab & " | PAX restante: " & paxRestante
        
        If paxRestante >= capacidadHab Then
            paxRestante = paxRestante - capacidadHab
        Else
            paxRestante = 0
        End If
        
        If paxRestante = 0 Then Exit For
    Next i
    
    Debug.Print "PAX restante final: " & paxRestante
    Debug.Print "Resultado: " & (paxRestante = 0)
    Debug.Print "=========================================="
    
    If paxRestante > 0 Then
        ReevaluacionViableGijonPorCapacidad = False
    Else
        ReevaluacionViableGijonPorCapacidad = True
    End If
End Function

Public Function FiltrarHabitacionesCompatiblesGijon( _
        ByVal unidadesRenuncia As Variant, _
        ByVal qSolicitado As Long, _
        ByVal rSolicitado As Long) As Variant
    
    Debug.Print "=== FiltrarHabitacionesCompatiblesGijon ==="
    Debug.Print "Q Solicitado (individuales): " & qSolicitado
    Debug.Print "R Solicitado (dobles): " & rSolicitado
    Debug.Print "Habitaciones recibidas: " & Join(unidadesRenuncia, ", ")
    
    Dim habCompatibles() As String
    Dim numCompatibles As Long
    Dim i As Long
    Dim tipoHab As String
    Dim numHab As String
    
    numCompatibles = 0
    
    ' Recorrer todas las habitaciones renunciadas
    For i = LBound(unidadesRenuncia) To UBound(unidadesRenuncia)
        numHab = Trim(CStr(unidadesRenuncia(i)))
        tipoHab = ClasificarHabitacionGijon(numHab)
        
        Debug.Print "Hab[" & i & "]: " & numHab & " -> Tipo: " & tipoHab
        
        Dim esCompatible As Boolean
        esCompatible = False
        
        ' L�GICA DE COMPATIBILIDAD:
        ' - Individual: solo compatible si solicita individuales (Q > 0)
        ' - Doble: compatible si solicita dobles (R > 0) O individuales (Q > 0)
        If tipoHab = "IND" Then
            If qSolicitado > 0 Then esCompatible = True
        ElseIf tipoHab = "DOB" Then
            If rSolicitado > 0 Or qSolicitado > 0 Then esCompatible = True
        End If
        
        Debug.Print "  �Es compatible?: " & esCompatible
        
        If esCompatible Then
            numCompatibles = numCompatibles + 1
            ReDim Preserve habCompatibles(1 To numCompatibles)
            habCompatibles(numCompatibles) = numHab
            Debug.Print "  >> AGREGADA a compatibles (total: " & numCompatibles & ")"
        End If
    Next i
    
    Debug.Print "Total compatibles encontradas: " & numCompatibles
    Debug.Print "Habitaciones compatibles: " & IIf(numCompatibles > 0, Join(habCompatibles, ", "), "ninguna")
    
    ' DEVOLVER TODAS LAS COMPATIBLES (sin filtrar por cantidad solicitada)
    ' La decisi�n de cu�ntas asignar la toma ReevaluarRenunciaResidencia
    If numCompatibles = 0 Then
        Debug.Print ">> DEVOLVIENDO EMPTY"
        FiltrarHabitacionesCompatiblesGijon = Empty
    Else
        Debug.Print ">> DEVOLVIENDO TODAS LAS COMPATIBLES"
        FiltrarHabitacionesCompatiblesGijon = habCompatibles
    End If
    Debug.Print "=========================================="
End Function

Public Function ClasificarHabitacionGijon(ByVal numHab As String) As String
    Select Case Trim(numHab)
        Case "1", "2", "3", "5", "6", "7", "OF.1", "OF.2", "OF.3"
            ClasificarHabitacionGijon = "DOB"
        Case "4", "EST.1", "EST.2", "EST.3"
            ClasificarHabitacionGijon = "IND"
        Case Else
            ClasificarHabitacionGijon = ""
    End Select
End Function

Public Sub ContarHabitacionesPorTipoGijon( _
        ByVal unidadesRenuncia As Variant, _
        ByRef numInd As Long, _
        ByRef numDob As Long)
    
    Debug.Print "=== ContarHabitacionesPorTipoGijon ==="
    
    Dim i As Long
    Dim tipoHab As String
    
    numInd = 0
    numDob = 0
    
    For i = LBound(unidadesRenuncia) To UBound(unidadesRenuncia)
        tipoHab = ClasificarHabitacionGijon(CStr(unidadesRenuncia(i)))
        Debug.Print "Hab: " & unidadesRenuncia(i) & " -> Tipo: " & tipoHab
        If tipoHab = "IND" Then
            numInd = numInd + 1
        ElseIf tipoHab = "DOB" Then
            numDob = numDob + 1
        End If
    Next i
    
    Debug.Print "Total IND: " & numInd & " | Total DOB: " & numDob
    Debug.Print "=========================================="
End Sub
' Verifica si una candidata puede ocupar AL MENOS las habitaciones que necesita
' de las disponibles en la renuncia, considerando compatibilidad de tipos
Public Function PuedeAdjudicarUnidadGijon(ws As Worksheet, _
                                          filaCandidata As Long, _
                                          unidadesRenuncia As Variant, _
                                          fechaEntrada As Date, _
                                          fechaSalida As Date, _
                                          filaExcluida As Long, _
                                          columnaUnidades As String, _
                                          qSolicitado As Long, _
                                          rSolicitado As Long) As Boolean
    
    Debug.Print "  === PuedeAdjudicarUnidadGijon ==="
    Debug.Print "  Q solicitado: " & qSolicitado & " | R solicitado: " & rSolicitado
    
    ' Obtener habitaciones compatibles con lo solicitado
    Dim habCompatibles As Variant
    habCompatibles = FiltrarHabitacionesCompatiblesGijon(unidadesRenuncia, qSolicitado, rSolicitado)
    
    If IsEmpty(habCompatibles) Then
        Debug.Print "  >> NO hay habitaciones compatibles"
        PuedeAdjudicarUnidadGijon = False
        Exit Function
    End If
    
    Debug.Print "  Habitaciones compatibles: " & Join(habCompatibles, ", ")
    
    ' Verificar disponibilidad de cada habitaci�n compatible
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    
    Dim habitacion As Variant
    Dim habitacionesDisponibles As Long
    habitacionesDisponibles = 0
    
    For Each habitacion In habCompatibles
        Dim estaDisponible As Boolean
        estaDisponible = True
        
        ' Buscar conflictos con otras reservas
        Dim i As Long, j As Long
        For i = 2 To ultimaFila
            If i <> filaCandidata And i <> filaExcluida Then
                Dim arrUnidad As Variant
                Dim valorUnidades As String
                valorUnidades = ws.Cells(i, columnaUnidades).Value
                
                If Trim(valorUnidades) <> "" Then
                    arrUnidad = Split(Replace(valorUnidades, " ", ""), ",")
                    
                    For j = LBound(arrUnidad) To UBound(arrUnidad)
                        If Trim(arrUnidad(j)) = Trim(habitacion) Then
                            ' Verificar solape de fechas
                            If HaySolapeReservas(ws.Cells(i, "L").Value, ws.Cells(i, "M").Value, fechaEntrada, fechaSalida) Then
                                estaDisponible = False
                                Debug.Print "  Hab " & habitacion & " ocupada por fila " & i
                                Exit For
                            End If
                        End If
                    Next j
                End If
            End If
            
            If Not estaDisponible Then Exit For
        Next i
        
        If estaDisponible Then
            habitacionesDisponibles = habitacionesDisponibles + 1
            Debug.Print "  Hab " & habitacion & " DISPONIBLE"
        End If
    Next habitacion
    
    ' Verificar si hay suficientes habitaciones disponibles
    Dim totalSolicitado As Long
    totalSolicitado = qSolicitado + rSolicitado
    
    Debug.Print "  Total habitaciones disponibles: " & habitacionesDisponibles
    Debug.Print "  Total habitaciones solicitadas: " & totalSolicitado
    
    If habitacionesDisponibles >= totalSolicitado Then
        Debug.Print "  >> PUEDE ADJUDICAR"
        PuedeAdjudicarUnidadGijon = True
    Else
        Debug.Print "  >> NO puede adjudicar (insuficientes disponibles)"
        PuedeAdjudicarUnidadGijon = False
    End If
    
    Debug.Print "  =================================="
End Function
