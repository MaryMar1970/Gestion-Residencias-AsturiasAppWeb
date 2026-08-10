Attribute VB_Name = "ReevaluacionSolicitudes"
' =========================================================================================
' Reevaluaci�n de solicitudes tras RENUNCIA en cualquier hoja RESIDENCIA (GEN�RICO)
' =========================================================================================
' Este m�dulo permite la reevaluaci�n y reasignaci�n de habitaciones o apartamentos tras una renuncia,
' funcionando para cualquier hoja de residencia: GIJ�N, OVIEDO, SOTO (u otras futuras).
'
' Par�metros personalizables:
'   - hojaNombre: nombre de la hoja de residencia (ej: "RESIDENCIA OVIEDO")
'   - filaRenuncia: fila donde se produce la renuncia
'   - columnaUnidades: columna donde se almacenan las habitaciones/apartamentos adjudicados (ej: "T" o "S")
'   - tipoResidencia: "GIJ�N", "OVIEDO", "SOTO" (seg�n la l�gica espec�fica de cada residencia)
'
' Para invocarlo desde el evento Worksheet_Change de cada hoja:
'   Call ReevaluarRenunciaResidencia("RESIDENCIA OVIEDO", Target.Row, "T", "OVIEDO")
'   Call ReevaluarRenunciaResidencia("RESIDENCIA GIJ�N", Target.Row, "S", "GIJ�N")
'   Call ReevaluarRenunciaResidencia("RESIDENCIA SOTO", Target.Row, "S", "SOTO")
'
' =========================================================================================

Option Explicit
' Obliga a declarar todas las variables, evitando errores por nombres mal escritos

'-------------------------------------------------------------------------------------------
' Funci�n: BuscarCandidatasValidas
' Prop�sito: Busca solicitudes que puedan ocupar las unidades liberadas por una renuncia
' Par�metros:
'   - ws: Hoja de trabajo donde buscar
'   - filaRenuncia: Fila de la solicitud que renuncia
'   - fechaEntradaRenuncia/Salida: Rango de fechas de la renuncia
'   - tipoResidencia: Tipo de residencia (ej. "SOTO")
'   - valorQ/R_Renuncia: Valores de las columnas Q y R de la renuncia
'   - unidadesRenuncia: Array con las unidades liberadas
'   - columnaUnidades: Letra de columna donde est�n las unidades (ej. "N")
' Retorno: Array con datos de candidatas v�lidas o Empty si no hay
'-------------------------------------------------------------------------------------------
Function BuscarCandidatasValidas(ws As Worksheet, _
    filaRenuncia As Long, _
    fechaEntradaRenuncia As Date, _
    fechaSalidaRenuncia As Date, _
    tipoResidencia As String, _
    valorQ_Renuncia As Variant, _
    valorR_Renuncia As Variant, _
    unidadesRenuncia As Variant, _
    columnaUnidades As String) As Variant
    
    Dim candidatas() As Variant
    Dim numCandidatas As Long: numCandidatas = 0
    Dim i As Long
    Dim ultimaFila As Long
    Dim colTelefono As String
    
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    colTelefono = IIf(UCase(tipoResidencia) = "OVIEDO", "W", "V")
    
    ' Control defensivo: si la celda de habitaciones est� vac�a, no hay nada que buscar
    If Trim(CStr(ws.Cells(filaRenuncia, columnaUnidades).Value)) = "" Then
       BuscarCandidatasValidas = Empty
       Exit Function
    End If
    
    For i = 2 To ultimaFila
        If i = filaRenuncia Then GoTo SiguienteI
        
        If UCase(Trim(ws.Cells(i, "P").Value)) = "NO" Then
            Dim fecEnt As Date, fecSal As Date
            fecEnt = ws.Cells(i, "L").Value
            fecSal = ws.Cells(i, "M").Value
            
            If fecEnt <= fechaSalidaRenuncia And fecSal >= fechaEntradaRenuncia Then
                Dim candidataValida As Boolean: candidataValida = True
                
                ' ========== VALIDACIONES ESPEC�FICAS POR RESIDENCIA ==========
                If UCase(tipoResidencia) = "OVIEDO" Then
                    Dim capRenuncia As Long
                    Dim capCandidata As Long
                    Dim supletoriasCandidata As Long
                    Dim paxCandidata As Long
                    Dim habIndCandidata As Long
                    Dim habDobCandidata As Long
                    Dim maxSupletoriasPermitidas As Long
                    
                    paxCandidata = val(ws.Cells(i, "O").Value)
                    habIndCandidata = val(ws.Cells(i, "Q").Value)
                    habDobCandidata = val(ws.Cells(i, "R").Value)
                    supletoriasCandidata = val(ws.Cells(i, "S").Value)
                    
                    capRenuncia = CapacidadTotalOviedo(valorQ_Renuncia, valorR_Renuncia, ws.Cells(filaRenuncia, "S").Value)
                    capCandidata = CapacidadTotalOviedo(habIndCandidata, habDobCandidata, supletoriasCandidata)
                    
                    If capCandidata > capRenuncia Then
                        candidataValida = False
                        GoTo SiguienteI
                    End If
                    
                    If paxCandidata <> capCandidata Then
                        candidataValida = False
                        GoTo SiguienteI
                    End If
                    
                    If supletoriasCandidata > 0 Then
                        If habDobCandidata = 0 Then
                            candidataValida = False
                            GoTo SiguienteI
                        End If
                        
                        Dim habitacionesDisponibles As String
                        habitacionesDisponibles = Trim(ws.Cells(filaRenuncia, "T").Value)
                        maxSupletoriasPermitidas = MaxSupletoriasPermitidasOviedo(habitacionesDisponibles)
                        
                        If supletoriasCandidata > maxSupletoriasPermitidas Then
                            candidataValida = False
                            GoTo SiguienteI
                        End If
                    End If
                    
                    Dim totalHabsDisponibles As Long
                    Dim totalHabsSolicitadas As Long
                    totalHabsDisponibles = valorQ_Renuncia + valorR_Renuncia
                    totalHabsSolicitadas = habIndCandidata + habDobCandidata
                    
                    If totalHabsSolicitadas > totalHabsDisponibles Then
                        candidataValida = False
                        GoTo SiguienteI
                    End If
                    
ElseIf UCase(tipoResidencia) = "GIJON" Then
    ' ========== VALIDACIONES PARA GIJ�N ==========
    Debug.Print vbCrLf & ">>> Evaluando fila " & i & " (N� ORDEN: " & ws.Cells(i, "A").Value & ")"
    Debug.Print "PAX: " & val(ws.Cells(i, "O").Value)
    Debug.Print "Q: " & val(ws.Cells(i, "Q").Value) & " | R: " & val(ws.Cells(i, "R").Value)
    
    Dim habLib() As String
    Dim arrTemp As Variant
    arrTemp = Split(Replace(ws.Cells(filaRenuncia, columnaUnidades).Value, " ", ""), ",")
    ReDim habLib(LBound(arrTemp) To UBound(arrTemp))
    
    Dim j As Long
    For j = LBound(arrTemp) To UBound(arrTemp)
        Select Case Trim(arrTemp(j))
            Case "1", "2", "3", "5", "6", "7", "OF.1", "OF.2", "OF.3"
                habLib(j) = "DOB"
            Case "4", "EST.1", "EST.2", "EST.3"
                habLib(j) = "IND"
            Case Else
                habLib(j) = ""
        End Select
    Next j
    
    Debug.Print "Habitaciones liberadas: " & Join(habLib, ", ")
    
    paxCandidata = val(ws.Cells(i, "O").Value)
    
    ' Validar viabilidad por capacidad
    If Not ReevaluacionViableGijonPorCapacidad(habLib, paxCandidata) Then
        Debug.Print ">>> RECHAZADA por capacidad"
        candidataValida = False
        GoTo SiguienteI
    End If
    
    ' NUEVA VALIDACI�N: Verificar compatibilidad de tipos de habitaci�n
    Dim qCandidata As Long, rCandidata As Long
    Dim numIndDisp As Long, numDobDisp As Long
    
    qCandidata = val(ws.Cells(i, "Q").Value)
    rCandidata = val(ws.Cells(i, "R").Value)
    
    Debug.Print "Contando habitaciones disponibles por tipo..."
    ' Contar habitaciones disponibles por tipo
    Call ContarHabitacionesPorTipoGijon(unidadesRenuncia, numIndDisp, numDobDisp)
    
    Debug.Print "Disponibles: " & numIndDisp & " IND, " & numDobDisp & " DOB"
    Debug.Print "Solicitadas: " & qCandidata & " IND, " & rCandidata & " DOB"
    
    ' Validar que hay suficientes habitaciones del tipo correcto
    ' Las dobles pueden usarse como individuales, pero no al rev�s
    If qCandidata > (numIndDisp + numDobDisp) Then
        Debug.Print ">>> RECHAZADA: solicita m�s individuales que disponibles (considerando degradaci�n)"
        candidataValida = False
        GoTo SiguienteI
    End If
    
    If rCandidata > numDobDisp Then
        Debug.Print ">>> RECHAZADA: solicita m�s dobles que disponibles"
        candidataValida = False
        GoTo SiguienteI
    End If
    
    Debug.Print ">>> APROBADA - continuando validaciones"
End If
                ' ========== FIN VALIDACIONES ESPEC�FICAS ==========
                
                If candidataValida Then
                    Dim fechaSolicitudCandidata As Date
                    fechaSolicitudCandidata = DateValue(ws.Cells(i, "B").Value)
                    Dim antiguedadDias As Long
                    antiguedadDias = Abs(DateDiff("d", fechaSolicitudCandidata, Date))
                    
                    If antiguedadDias <= 11 Then
                        ' Usar funci�n espec�fica para GIJON
                        Dim puedeAdjudicar As Boolean
                        If UCase(tipoResidencia) = "GIJON" Then
                            puedeAdjudicar = PuedeAdjudicarUnidadGijon(ws, i, unidadesRenuncia, fecEnt, fecSal, filaRenuncia, columnaUnidades, qCandidata, rCandidata)
                        Else
                            puedeAdjudicar = PuedeAdjudicarUnidad(ws, i, unidadesRenuncia, fecEnt, fecSal, filaRenuncia, columnaUnidades)
                        End If
                        
                        If puedeAdjudicar Then
                        Dim usaSupletorias As Boolean
                            usaSupletorias = False
                            If UCase(tipoResidencia) = "OVIEDO" Then
                                If val(ws.Cells(i, "S").Value) > 0 Then usaSupletorias = True
                            End If
                            
                            numCandidatas = numCandidatas + 1
                            ReDim Preserve candidatas(1 To 14, 1 To numCandidatas)
                            candidatas(1, numCandidatas) = i
                            candidatas(2, numCandidatas) = ws.Cells(i, "B").Value
                            candidatas(3, numCandidatas) = antiguedadDias
                            candidatas(4, numCandidatas) = ws.Cells(i, "A").Value
                            candidatas(5, numCandidatas) = ws.Cells(i, "P").Address
                            candidatas(6, numCandidatas) = Format(ws.Cells(i, "L").Value, "dd/mm/yyyy") _
                                & " - " & Format(ws.Cells(i, "M").Value, "dd/mm/yyyy")
                            candidatas(7, numCandidatas) = ws.Cells(i, "K").Value
                            candidatas(8, numCandidatas) = ws.Cells(i, colTelefono).Value
                            candidatas(9, numCandidatas) = IIf(i < filaRenuncia, "ANTERIOR", "POSTERIOR")
                            candidatas(10, numCandidatas) = usaSupletorias
                            candidatas(11, numCandidatas) = val(ws.Cells(i, "O").Value)
                            candidatas(12, numCandidatas) = val(ws.Cells(i, "Q").Value)
                            candidatas(13, numCandidatas) = val(ws.Cells(i, "R").Value)
                            candidatas(14, numCandidatas) = val(ws.Cells(i, "S").Value)
                        End If
                    End If
                End If
            End If
        End If
SiguienteI:
    Next i
    
    If numCandidatas = 0 Then
        BuscarCandidatasValidas = Empty
    Else
        BuscarCandidatasValidas = candidatas
    End If
End Function

'-------------------------------------------------------------------------------------------
' Sub: ReevaluarRenunciaResidencia
' Prop�sito: Gestiona el proceso completo de reevaluaci�n de una renuncia
' Par�metros:
'   - hojaNombre: Nombre de la hoja de trabajo
'   - filaRenuncia: Fila de la solicitud que renuncia
'   - columnaUnidades: Columna donde est�n las unidades (ej. "N")
'   - tipoResidencia: Tipo de residencia (ej. "SOTO")
'-------------------------------------------------------------------------------------------
Public Sub ReevaluarRenunciaResidencia(ByVal hojaNombre As String, ByVal filaRenuncia As Long, ByVal columnaUnidades As String, ByVal tipoResidencia As String)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(hojaNombre)
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    If filaRenuncia < 2 Or filaRenuncia > ultimaFila Then Exit Sub
    
    ' Control defensivo: si la celda de habitaciones est� vac�a, no se puede realizar reevaluaci�n por renuncia
    If Trim(CStr(ws.Cells(filaRenuncia, columnaUnidades).Value)) = "" Then
        MsgBox "No hay habitaciones asignadas en esta fila para reasignar por renuncia.", vbExclamation, "Reasignaci�n no disponible"
        Exit Sub
    End If
    
    Dim fechaEntradaRenuncia As Date, fechaSalidaRenuncia As Date
    Dim unidadesRenuncia As Variant
    Dim valorQ_Renuncia As Variant, valorR_Renuncia As Variant
    Dim fechaSolicitudRenuncia As Date
    Dim valorUnidadesRenuncia As String
    
    fechaEntradaRenuncia = ws.Cells(filaRenuncia, "L").Value
    fechaSalidaRenuncia = ws.Cells(filaRenuncia, "M").Value
    unidadesRenuncia = Split(Replace(ws.Cells(filaRenuncia, columnaUnidades).Value, " ", ""), ",")
    valorQ_Renuncia = ws.Cells(filaRenuncia, "Q").Value
    valorR_Renuncia = ws.Cells(filaRenuncia, "R").Value
    fechaSolicitudRenuncia = DateValue(ws.Cells(filaRenuncia, "B").Value)
    valorUnidadesRenuncia = ws.Cells(filaRenuncia, columnaUnidades).Value
    
    Dim candidatas As Variant
    candidatas = BuscarCandidatasValidas(ws, filaRenuncia, fechaEntradaRenuncia, fechaSalidaRenuncia, tipoResidencia, valorQ_Renuncia, valorR_Renuncia, unidadesRenuncia, columnaUnidades)
    
    Dim numCandidatas As Long
    If IsEmpty(candidatas) Then
        numCandidatas = 0
    Else
        numCandidatas = UBound(candidatas, 2)
    End If
    
    ' Ordenar candidatas
    If numCandidatas > 1 Then
        Dim j As Long, k As Long, t As Long
        Dim tmpVal As Variant
        For j = 1 To numCandidatas - 1
            For k = j + 1 To numCandidatas
                Dim swap As Boolean
                swap = False
                
                If UCase(tipoResidencia) = "OVIEDO" Then
                    If candidatas(10, j) = True And candidatas(10, k) = False Then
                        swap = True
                    End If
                End If
                
                If Not swap Then
                    If candidatas(3, j) < candidatas(3, k) Then
                        swap = True
                    End If
                End If
                
                If swap Then
                    For t = 1 To UBound(candidatas, 1)
                        tmpVal = candidatas(t, j)
                        candidatas(t, j) = candidatas(t, k)
                        candidatas(t, k) = tmpVal
                    Next t
                End If
            Next k
        Next j
    End If
    
            If numCandidatas > 0 Then
        Dim mostrarHabs As String
        Dim numHabs As Long
        numHabs = UBound(unidadesRenuncia) - LBound(unidadesRenuncia) + 1
        If numHabs > 5 Then
            Dim tempHabs() As String
            ReDim tempHabs(0 To 4)
            Dim hIdx As Long
            For hIdx = 0 To 4
                tempHabs(hIdx) = CStr(unidadesRenuncia(LBound(unidadesRenuncia) + hIdx))
            Next hIdx
            mostrarHabs = Join(tempHabs, ", ") & " ... (y " & (numHabs - 5) & " m�s)"
        Else
            mostrarHabs = Join(unidadesRenuncia, ", ")
        End If

        Dim seleccion As String
        Dim idx As Long
        Dim pStart As Long: pStart = 1
        Dim pEnd As Long
        Dim pageSize As Long: pageSize = 5
        
        Do
            Dim opciones As String: opciones = ""
            pEnd = pStart + pageSize - 1
            If pEnd > numCandidatas Then pEnd = numCandidatas
            
            Dim idxCand As Long
            For idxCand = pStart To pEnd
                Dim infoAdicional As String
                If UCase(tipoResidencia) = "OVIEDO" Then
                    infoAdicional = "   Pax: " & candidatas(11, idxCand) & _
                                    " | Ind.: " & candidatas(12, idxCand) & _
                                    " | Doble: " & candidatas(13, idxCand) & _
                                    " | Suple.: " & candidatas(14, idxCand)
                ElseIf UCase(tipoResidencia) = "GIJON" Then
                    infoAdicional = "   Pax: " & candidatas(11, idxCand) & _
                                    " | Ind.: " & candidatas(12, idxCand) & _
                                    " | Doble: " & candidatas(13, idxCand)
                Else
                    infoAdicional = "   Pax: " & candidatas(11, idxCand)
                End If

                opciones = opciones & "[ " & idxCand & " ] N� ORDEN: " & candidatas(4, idxCand) & _
                    " (" & candidatas(9, idxCand) & ")" & vbCrLf & _
                    "   Fecha solicitud: " & Format(candidatas(2, idxCand), "dd/mm/yyyy") & " | Tel�fono: " & candidatas(8, idxCand) & vbCrLf & _
                    "   Nombre: " & candidatas(7, idxCand) & vbCrLf & _
                    "   Entrada: " & Split(candidatas(6, idxCand), " - ")(0) & " - Salida: " & Split(candidatas(6, idxCand), " - ")(1) & vbCrLf & _
                    infoAdicional
                If idxCand < pEnd Then opciones = opciones & vbCrLf & vbCrLf Else opciones = opciones & vbCrLf
            Next idxCand
            
            Dim pagInstructions As String: pagInstructions = ""
            If pEnd < numCandidatas Then
                pagInstructions = pagInstructions & "[ S ] Ver siguientes candidatos..." & vbCrLf
            End If
            If pStart > 1 Then
                pagInstructions = pagInstructions & "[ A ] Ver anteriores candidatos..." & vbCrLf
            End If
            
            If pagInstructions <> "" Then
                opciones = opciones & vbCrLf & "Paginaci�n:" & vbCrLf & pagInstructions
            End If
            
            seleccion = InputBox("Habitaci�n disponible tras renuncia: N� " & _
                mostrarHabs & vbCrLf & vbCrLf & opciones & vbCrLf & _
                "Ingrese el [ n� ] de la solicitud a reevaluar:", _
                "REEVALUACI�N SOLICITUDES POR RENUNCIA")
            
            If seleccion = "" Then Exit Sub
            
            seleccion = Trim(UCase(seleccion))
            If seleccion = "S" And pEnd < numCandidatas Then
                pStart = pStart + pageSize
            ElseIf seleccion = "A" And pStart > 1 Then
                pStart = pStart - pageSize
                If pStart < 1 Then pStart = 1
            ElseIf IsNumeric(seleccion) Then
                idx = CLng(seleccion)
                If idx >= 1 And idx <= numCandidatas Then Exit Do
                MsgBox "Por favor, introduzca un n� de opci�n v�lido (entre 1 y " & numCandidatas & ").", vbExclamation
            Else
                MsgBox "Entrada no reconocida. Ingrese un n�mero de opci�n, ""S"" para ver m�s o ""A"" para volver atr�s.", vbExclamation
            End If
        Loop
        
        Dim filaCandidata As Long: filaCandidata = candidatas(1, idx)
        Dim valorQ_Candidata As Variant, valorR_Candidata As Variant
        valorQ_Candidata = ws.Cells(filaCandidata, "Q").Value
        valorR_Candidata = ws.Cells(filaCandidata, "R").Value
        
        Dim palabraUnidad As String
        If UCase(tipoResidencia) = "SOTO" Then
            palabraUnidad = "Ap."
        Else
            palabraUnidad = "Hab."
        End If
        
        ' ===== DEBUG CR�TICO =====
        Debug.Print vbCrLf & "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@"
        Debug.Print "LLEGANDO A ZONA DE PROCESAMIENTO"
        Debug.Print "tipoResidencia: [" & tipoResidencia & "]"
        Debug.Print "UCase(tipoResidencia): [" & UCase(tipoResidencia) & "]"
        Debug.Print "�Es GIJON?: " & (UCase(tipoResidencia) = "GIJON")
        Debug.Print "valorQ_Candidata: " & valorQ_Candidata
        Debug.Print "valorR_Candidata: " & valorR_Candidata
        Debug.Print "columnaUnidades: " & columnaUnidades
        Debug.Print "@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@" & vbCrLf
        
       ' ========== PROCESAMIENTO EXCLUSIVO PARA GIJON ==========
        If UCase(tipoResidencia) = "GIJON" Then
            Debug.Print vbCrLf & "############################################"
            Debug.Print "### INICIANDO PROCESAMIENTO GIJON ###"
            Debug.Print "Fila Candidata: " & filaCandidata
            Debug.Print "N� ORDEN Candidata: " & candidatas(4, idx)
            Debug.Print "Q Candidata: " & valorQ_Candidata
            Debug.Print "R Candidata: " & valorR_Candidata
            Debug.Print "Habitaciones renunciadas: " & Join(unidadesRenuncia, ", ")
            Debug.Print "############################################" & vbCrLf
            
            Dim habCompatibles As Variant
            habCompatibles = FiltrarHabitacionesCompatiblesGijon(unidadesRenuncia, valorQ_Candidata, valorR_Candidata)
            
            Debug.Print vbCrLf & ">>> RESULTADO FiltrarHabitacionesCompatiblesGijon:"
            If IsEmpty(habCompatibles) Then
                Debug.Print ">>> EMPTY (ninguna compatible)"
            Else
                Debug.Print ">>> Habitaciones compatibles: " & Join(habCompatibles, ", ")
                Debug.Print ">>> Cantidad: " & (UBound(habCompatibles) - LBound(habCompatibles) + 1)
            End If
            Debug.Print ""
            
            If IsEmpty(habCompatibles) Then
                MsgBox "Error: No hay habitaciones compatibles disponibles.", vbCritical
                Exit Sub
            End If
            
            Dim numHabCompatibles As Long
            numHabCompatibles = UBound(habCompatibles) - LBound(habCompatibles) + 1
            Dim numHabSolicitadas As Long
            numHabSolicitadas = valorQ_Candidata + valorR_Candidata
            
            Debug.Print "N�mero habitaciones compatibles: " & numHabCompatibles
            Debug.Print "N�mero habitaciones solicitadas: " & numHabSolicitadas
            
            ' CASO A: Solicita MENOS habitaciones que las compatibles disponibles
            ' -> Debe elegir cu�l(es) asignar
            If numHabSolicitadas < numHabCompatibles Then
                Debug.Print ">>> CASO A: Solicita MENOS que disponibles. Mostrando men� de selecci�n..."
                
                ' Si solicita solo 1 habitaci�n, mostrar men� simple
                If numHabSolicitadas = 1 Then
                    Dim unidadOpciones As String, h As Integer
                    unidadOpciones = ""
                    For h = LBound(habCompatibles) To UBound(habCompatibles)
                        unidadOpciones = unidadOpciones & "[ " & (h - LBound(habCompatibles) + 1) & " ] " & palabraUnidad & " " & habCompatibles(h) & vbCrLf & vbCrLf
                    Next h
                    
                    Dim unidadSeleccion As String
                    Dim idxUnidad As Integer
                    Do
                        unidadSeleccion = InputBox("Ingrese, de la lista, el [ n� ] de la habitaci�n (" & palabraUnidad & "):" & vbCrLf & vbCrLf & unidadOpciones, "SELECCI�N HABITACI�N A ADJUDICAR")
                        If unidadSeleccion = "" Then Exit Sub
                        If IsNumeric(unidadSeleccion) Then
                            idxUnidad = CInt(unidadSeleccion)
                            If idxUnidad >= 1 And idxUnidad <= numHabCompatibles Then Exit Do
                        End If
                        MsgBox "Por favor, introduzca un n�mero v�lido de la lista de unidades.", vbExclamation
                    Loop
                    
                    Debug.Print "Habitaci�n seleccionada: " & habCompatibles(idxUnidad - 1 + LBound(habCompatibles))
                    
                    ws.Range(candidatas(5, idx)).Value = "REEVALUADA"
                    ws.Cells(filaCandidata, columnaUnidades).Value = habCompatibles(idxUnidad - 1 + LBound(habCompatibles))
                    
                    Debug.Print "ADJUDICACI�N COMPLETADA (CASO A - 1 hab)"
                    Debug.Print "############################################" & vbCrLf
                    
                    MsgBox "Reasignada " & palabraUnidad & " " & habCompatibles(idxUnidad - 1 + LBound(habCompatibles)) & _
                        " a la solicitud con N� ORDEN: " & candidatas(4, idx), vbInformation
                    Exit Sub
                Else
                    ' Solicita m�ltiples habitaciones pero menos que las disponibles
                    ' Seleccionar las primeras N habitaciones compatibles seg�n el tipo
                    Dim habSeleccionadas() As String
                    Dim numSeleccionadas As Long
                    Dim doblesAsignadas As Long
                    Dim individualesAsignadas As Long
                    
                    numSeleccionadas = 0
                    doblesAsignadas = 0
                    individualesAsignadas = 0
                    
                    ' Primera pasada: asignar dobles si se solicitan
                    If valorR_Candidata > 0 Then
                        For j = LBound(habCompatibles) To UBound(habCompatibles)
                            If doblesAsignadas >= valorR_Candidata Then Exit For
                            If ClasificarHabitacionGijon(habCompatibles(j)) = "DOB" Then
                                numSeleccionadas = numSeleccionadas + 1
                                ReDim Preserve habSeleccionadas(1 To numSeleccionadas)
                                habSeleccionadas(numSeleccionadas) = habCompatibles(j)
                                doblesAsignadas = doblesAsignadas + 1
                            End If
                        Next j
                    End If
                    
                    ' Segunda pasada: asignar individuales si se solicitan
                    If valorQ_Candidata > 0 Then
                        For j = LBound(habCompatibles) To UBound(habCompatibles)
                            If individualesAsignadas >= valorQ_Candidata Then Exit For
                            
                            ' Verificar si ya fue asignada
                            Dim yaAsignada As Boolean
                            yaAsignada = False
                            For k = 1 To numSeleccionadas
                                If habSeleccionadas(k) = habCompatibles(j) Then
                                    yaAsignada = True
                                    Exit For
                                End If
                            Next k
                            
                            If Not yaAsignada And ClasificarHabitacionGijon(habCompatibles(j)) = "IND" Then
                                numSeleccionadas = numSeleccionadas + 1
                                ReDim Preserve habSeleccionadas(1 To numSeleccionadas)
                                habSeleccionadas(numSeleccionadas) = habCompatibles(j)
                                individualesAsignadas = individualesAsignadas + 1
                            End If
                        Next j
                    End If
                    
                    Dim habitacionesAdjudicar As String
                    habitacionesAdjudicar = Join(habSeleccionadas, ", ")
                    
                    Debug.Print "Habitaciones seleccionadas autom�ticamente: " & habitacionesAdjudicar
                    
                    ws.Range(candidatas(5, idx)).Value = "REEVALUADA"
                    ws.Cells(filaCandidata, columnaUnidades).Value = habitacionesAdjudicar
                    
                    ' Mensaje de confirmaci�n
                    Dim arrUnidad As Variant
                    arrUnidad = habSeleccionadas
                    Dim unidadMsg As String
                    Dim nUnidad As Integer
                    nUnidad = UBound(arrUnidad) - LBound(arrUnidad) + 1
                    Dim m As Integer
                    Select Case nUnidad
                        Case 1
                            unidadMsg = palabraUnidad & " " & arrUnidad(LBound(arrUnidad))
                        Case 2
                            unidadMsg = palabraUnidad & " " & arrUnidad(LBound(arrUnidad)) & " y " & palabraUnidad & " " & arrUnidad(UBound(arrUnidad))
                        Case Else
                            For m = LBound(arrUnidad) To UBound(arrUnidad) - 1
                                unidadMsg = unidadMsg & palabraUnidad & " " & arrUnidad(m) & ", "
                            Next m
                            unidadMsg = unidadMsg & "y " & palabraUnidad & " " & arrUnidad(UBound(arrUnidad))
                    End Select
                    
                    Debug.Print "ADJUDICACI�N COMPLETADA (CASO A - m�ltiples)"
                    Debug.Print "############################################" & vbCrLf
                    
                    MsgBox "Reasignadas las " & unidadMsg & " a la solicitud con N� ORDEN: " & candidatas(4, idx), vbInformation
                    Exit Sub
                End If
            
            ' CASO B: Solicita EXACTAMENTE las habitaciones compatibles disponibles
            ' -> Adjudicaci�n autom�tica de todas
            Else
                Debug.Print ">>> CASO B: Adjudicaci�n autom�tica (solicita todas las compatibles)"
                
                habitacionesAdjudicar = Join(habCompatibles, ", ")
                
                Debug.Print "Habitaciones a adjudicar: " & habitacionesAdjudicar
                Debug.Print "Escribiendo en celda " & columnaUnidades & filaCandidata & ": " & habitacionesAdjudicar
                
                ws.Range(candidatas(5, idx)).Value = "REEVALUADA"
                ws.Cells(filaCandidata, columnaUnidades).Value = habitacionesAdjudicar
                
                Debug.Print "Valor escrito en hoja: " & ws.Cells(filaCandidata, columnaUnidades).Value
                
                ' Mensaje de confirmaci�n
                arrUnidad = habCompatibles
                nUnidad = UBound(arrUnidad) - LBound(arrUnidad) + 1
                Select Case nUnidad
                    Case 1
                        unidadMsg = palabraUnidad & " " & arrUnidad(LBound(arrUnidad))
                    Case 2
                        unidadMsg = palabraUnidad & " " & arrUnidad(LBound(arrUnidad)) & " y " & palabraUnidad & " " & arrUnidad(UBound(arrUnidad))
                    Case Else
                        For m = LBound(arrUnidad) To UBound(arrUnidad) - 1
                            unidadMsg = unidadMsg & palabraUnidad & " " & arrUnidad(m) & ", "
                        Next m
                        unidadMsg = unidadMsg & "y " & palabraUnidad & " " & arrUnidad(UBound(arrUnidad))
                End Select
                
                Debug.Print "ADJUDICACI�N COMPLETADA (CASO B)"
                Debug.Print "############################################" & vbCrLf
                
                MsgBox "Reasignadas las " & unidadMsg & " a la solicitud con N� ORDEN: " & candidatas(4, idx), vbInformation
                Exit Sub
            End If
        End If
        ' ========== FIN PROCESAMIENTO GIJON ==========        ' ========== PROCESAMIENTO PARA OVIEDO Y SOTO (SIN CAMBIOS) ==========
        If (valorQ_Renuncia + valorR_Renuncia > 1) And (valorQ_Candidata + valorR_Candidata = 1) Then
            If UBound(unidadesRenuncia) > 0 Then
                unidadOpciones = ""
                For h = LBound(unidadesRenuncia) To UBound(unidadesRenuncia)
                    unidadOpciones = unidadOpciones & "[ " & (h - LBound(unidadesRenuncia) + 1) & " ] " & palabraUnidad & " " & unidadesRenuncia(h) & vbCrLf & vbCrLf
                Next h
                
                Do
                    unidadSeleccion = InputBox("Ingrese, de la lista, el [ n� ] de la habitaci�n (" & palabraUnidad & "):" & vbCrLf & vbCrLf & unidadOpciones, "SELECCI�N HABITACI�N A ADJUDICAR")
                    If unidadSeleccion = "" Then Exit Sub
                    If IsNumeric(unidadSeleccion) Then
                        idxUnidad = CInt(unidadSeleccion)
                        If idxUnidad >= 1 And idxUnidad <= UBound(unidadesRenuncia) - LBound(unidadesRenuncia) + 1 Then Exit Do
                    End If
                    MsgBox "Por favor, introduzca un n�mero v�lido de la lista de unidades.", vbExclamation
                Loop
                
                ws.Range(candidatas(5, idx)).Value = "REEVALUADA"
                ws.Cells(filaCandidata, columnaUnidades).Value = unidadesRenuncia(idxUnidad - 1 + LBound(unidadesRenuncia))
                MsgBox "Reasignada " & palabraUnidad & " " & unidadesRenuncia(idxUnidad - 1 + LBound(unidadesRenuncia)) & _
                    " a la solicitud con N� ORDEN: " & candidatas(4, idx), vbInformation
                Exit Sub
            End If
        End If
        
        ws.Range(candidatas(5, idx)).Value = "REEVALUADA"
        ws.Cells(filaCandidata, columnaUnidades).Value = valorUnidadesRenuncia
        
        If UCase(tipoResidencia) = "OVIEDO" Then
            Dim qFinal As Long, rFinal As Long, sFinal As Long
            qFinal = val(ws.Cells(filaCandidata, "Q").Value)
            rFinal = val(ws.Cells(filaCandidata, "R").Value)
            sFinal = val(ws.Cells(filaCandidata, "S").Value)
            
            If sFinal > 0 And rFinal = 0 Then
                ws.Cells(filaCandidata, "S").ClearContents
            End If
        End If
        
        arrUnidad = Split(Replace(valorUnidadesRenuncia, " ", ""), ",")
        unidadMsg = ""
        nUnidad = UBound(arrUnidad) - LBound(arrUnidad) + 1
        Select Case nUnidad
            Case 1
                unidadMsg = palabraUnidad & " " & arrUnidad(0)
            Case 2
                unidadMsg = palabraUnidad & " " & arrUnidad(0) & " y " & palabraUnidad & " " & arrUnidad(1)
            Case Else
                For m = LBound(arrUnidad) To UBound(arrUnidad) - 1
                    unidadMsg = unidadMsg & palabraUnidad & " " & arrUnidad(m) & ", "
                Next m
                unidadMsg = unidadMsg & "y " & palabraUnidad & " " & arrUnidad(UBound(arrUnidad))
        End Select
        MsgBox "Reasignadas las " & unidadMsg & " a la solicitud con N� ORDEN: " & candidatas(4, idx), vbInformation
    Else
        MsgBox "No se encontraron candidatas v�lidas para la/s habitaci�n/es renunciada/s en " & hojaNombre & ".", vbInformation, "Reevaluaci�n"
    End If
End Sub

'-------------------------------------------------------------------------------------------
' Funci�n: PuedeAdjudicarUnidad
' Prop�sito: Verifica si las unidades pueden adjudicarse sin conflictos
' Par�metros:
'   - ws: Hoja de trabajo
'   - filaCandidata: Fila de la candidata
'   - unidadesRenuncia: Unidades a adjudicar
'   - fechaEntrada/Salida: Rango de fechas de la candidata
'   - filaExcluida: Fila a ignorar (normalmente la renuncia)
'   - columnaUnidades: Columna donde buscar unidades
' Retorno: True si puede adjudicar sin conflictos
'-------------------------------------------------------------------------------------------
Private Function PuedeAdjudicarUnidad(ws As Worksheet, filaCandidata As Long, unidadesRenuncia As Variant, _
                                          fechaEntrada As Date, fechaSalida As Date, filaExcluida As Long, columnaUnidades As String) As Boolean
    Dim ultimaFila As Long: ultimaFila = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
    Dim i As Long, j As Long
    Dim habitacion As Variant
    PuedeAdjudicarUnidad = True  ' Asume v�lido hasta encontrar conflicto

    ' Para cada unidad en la renuncia
    For Each habitacion In unidadesRenuncia
        ' Busca en todas las filas
        For i = 2 To ultimaFila
            ' Ignora fila candidata y renuncia
            If i <> filaCandidata And i <> filaExcluida Then
                Dim arrUnidad As Variant
                arrUnidad = Split(Replace(ws.Cells(i, columnaUnidades).Value, " ", ""), ",")
                ' Verifica cada unidad en la fila
                For j = LBound(arrUnidad) To UBound(arrUnidad)
                    If Trim(arrUnidad(j)) = Trim(habitacion) Then
                        ' Si hay solape de fechas, no es v�lido
                        If HaySolapeReservas(ws.Cells(i, "L").Value, ws.Cells(i, "M").Value, fechaEntrada, fechaSalida) Then
                            PuedeAdjudicarUnidad = False
                            Exit Function
                        End If
                    End If
                Next j
            End If
        Next i
    Next habitacion
End Function

'-------------------------------------------------------------------------------------------
' Funci�n: HaySolapeReservas
' Prop�sito: Determina si dos rangos de fechas se solapan
' Par�metros: Fechas de inicio/fin de ambos rangos
' Retorno: True si hay solape
'-------------------------------------------------------------------------------------------
Public Function HaySolapeReservas(inicio1 As Date, fin1 As Date, inicio2 As Date, fin2 As Date) As Boolean
    HaySolapeReservas = Not (fin1 <= inicio2 Or inicio1 >= fin2)
End Function

'-------------------------------------------------------------------------------------------
' Funci�n: CapacidadTotalOviedo
' Prop�sito:
'   Calcula la capacidad total de una adjudicaci�n en RESIDENCIA OVIEDO,
'   teniendo en cuenta habitaciones individuales, dobles y camas supletorias.
'
' Reglas de negocio:
'   - Hab. individual (Q) = 1 PAX
'   - Hab. doble (R) = 2 PAX
'   - Cama supletoria (S) = 1 PAX
'   - Las supletorias SIEMPRE van ligadas a dobles (R >= 1)
'
' Nota:
'   La validaci�n del m�ximo permitido por habitaci�n ya se controla
'   en los eventos de la hoja (Validaci�n de camas supletorias).
'-------------------------------------------------------------------------------------------
Public Function CapacidadTotalOviedo( _
    ByVal numHabInd As Long, _
    ByVal numHabDob As Long, _
    ByVal numCamasSuplet As Long) As Long

    Dim capacidad As Long

    ' Capacidad base por habitaciones
    capacidad = (numHabInd * 1) + (numHabDob * 2)

    ' A�ade camas supletorias solo si hay dobles
    If numHabDob > 0 And numCamasSuplet > 0 Then
        capacidad = capacidad + numCamasSuplet
    End If

    CapacidadTotalOviedo = capacidad
End Function

'-------------------------------------------------------------------------------------------
' Funci�n: MaxSupletoriasPermitidasOviedo
' Prop�sito: Calcula el m�ximo de camas supletorias f�sicamente posibles
'            seg�n las habitaciones espec�ficas asignadas
' Par�metros:
'   - habitaciones: String con habitaciones separadas por comas (ej: "13,14")
' Retorno: N�mero m�ximo de supletorias permitidas
'-------------------------------------------------------------------------------------------
Private Function MaxSupletoriasPermitidasOviedo(ByVal habitaciones As String) As Long
    ' Retorna el m�ximo de camas supletorias f�sicamente posibles
    ' seg�n las habitaciones espec�ficas asignadas en OVIEDO
    '
    ' Reglas OVIEDO (seg�n tu caso):
    '   Hab 1, 2, 3  ? admiten 1 supletoria cada una
    '   Hab 12       ? admite 2 supletorias
    '   Hab 13, 14   ? NO admiten supletorias
    '   Hab 6        ? NO admite supletorias (individual)
    '   Otras        ? NO admiten supletorias por defecto
    
    If Trim(habitaciones) = "" Then
        MaxSupletoriasPermitidasOviedo = 0
        Exit Function
    End If
    
    Dim arr As Variant
    Dim hab As Variant
    Dim maxSuplet As Long
    
    arr = Split(Replace(habitaciones, " ", ""), ",")
    maxSuplet = 0
    
    For Each hab In arr
        Select Case Trim(CStr(hab))
            Case "1", "2", "3"
                maxSuplet = maxSuplet + 1
            Case "12"
                maxSuplet = maxSuplet + 2
            Case Else
                ' Habitaciones 6, 13, 14 y otras NO admiten supletorias
                maxSuplet = maxSuplet + 0
        End Select
    Next hab
    
    MaxSupletoriasPermitidasOviedo = maxSuplet
End Function
