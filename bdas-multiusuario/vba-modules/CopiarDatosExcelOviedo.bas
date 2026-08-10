Attribute VB_Name = "CopiarDatosExcelOviedo"
' ************************************************************
' M�dulo: CopiaDatosExcelOviedo
' Descripci�n:
'   - Copia datos desde "REG-UNIDAD 2025" a "RESIDENCIA OVIEDO".
'   - Gestiona columnas especiales (G/D/E/F).
'   - Maneja fechas correctamente preservando formato.
'   - Borra la columna T seg�n reglas de resoluci�n.
'   - Al final, replica autom�ticamente la columna A en AC.
'   - Marca "SI" en AB si columna C tiene valor num�rico.
' ************************************************************

Sub CopiaDatosExcelOviedo()
    Dim wbOrigen As Workbook, wsOrigen As Worksheet
    Dim wbDestino As Workbook, wsDestino As Worksheet
    Dim rutaOrigen As Variant
    Dim columnasMapeo As Variant, columnasDestino As Variant
    Dim columnasFormulas() As Boolean
    Dim ultimaFilaOrigen As Long, numFilas As Long
    Dim i As Long, idx As Long
    Dim abiertoPorMacro As Boolean
    Dim calcPrevio As XlCalculation
    Dim eventosPrevios As Boolean, pantallaPrevio As Boolean

    ' Variables de uso en bloque 9
    Dim arrFechas As Variant
    Dim r As Long
    Dim colONum As Long, colDNum As Long

    ' --- Reglas y listas para casos especiales G/D/E/F ---
    Dim listaG As Variant, listaD As Variant, listaE As Variant, listaF As Variant
    Dim dictG As Object, posG As Variant

    ' Inicializar listas (orden 1:1)
    listaG = Array("1, 1, 1", "1, 1, 2", "1, 1, 2", "1, 1, 3", "1, 1, 4", "2, 8, 1", _
                   "1, 2, 2", "1, 3, 2", "1, 3, 2", "1, 4, 2", "1, 4, 2", _
                   "2, 1, 1", "2, 1, 2", "2, 1, 3", "2, 1, 4", "2, 1, 5", _
                   "2, 1, 6", "2, 1, 7", "2, 1, 8", "2, 2, 7", "2, 3, 7", _
                   "2, 4, 7", "2, 5, 7", "2, 6, 7", "2, 7, 7", "2, 8, 7", _
                   "2, 9, 7", "2, 9, 7", "2, 10, 7", "2, 10, 7", "2, 10, 7")
    listaD = Array("Comisi�n NO indem.", "Comisi�n NO indem.", "Comisi�n NO indem.", "Destino", "Comisi�n", "Enfermedad", _
                   "Comisi�n", "Comisi�n", "Destino", "Comisi�n", "Destino", _
                   "Enfermedad", "Otros", "Urgencia", "Sepelio", "M�x. Estancia", _
                   "Otros", "Otros", "Otros", "Otros", "Otros", _
                   "Otros", "Otros", "Otros", "Otros", "Otros", _
                   "Otros", "Otros", "Otros", "Otros", "Otros")
    listaE = Array("GC", "GC", "GC", "GC", "GC", "GC", _
                   "Alumno", "Militar en GC", "Militar en GC", "Funcionario en GC", "Funcionario en GC", _
                   "GC", "GC", "GC", "GC", "GC", _
                   "GC", "GC", "GC", "Alumno", "GC", _
                   "Militar en GC", "Funcionario en GC", "GC", "GC", "GC", _
                   "GC", "GC", "Militar no GC", "Militar no GC", "Militar no GC")
    listaF = Array("Viogen", "Activo", "Reserva activo", "Activo", "Activo", "Retirado", _
                   "Activo", "Activo", "Activo", "Activo", "Activo", _
                   "Activo", "Viogen", "Activo", "Activo", "Activo", _
                   "Asociaci�n", "Activo", "Reserva activo", "Activo", "Reserva", _
                   "Activo", "Activo", "Excedencia", "Especiales", "Retirado", _
                   "Viuda", "Huerfano", "Activo", "Reserva", "Retirado")

    Set dictG = CreateObject("Scripting.Dictionary")
    For i = LBound(listaG) To UBound(listaG)
        dictG(Replace(Trim(listaG(i)), " ", "")) = i
    Next i

    ' --- 1) Validaci�n hoja destino ---
    Set wbDestino = ThisWorkbook
    On Error Resume Next
    Set wsDestino = wbDestino.Worksheets("RESIDENCIA OVIEDO")
    On Error GoTo 0
    If wsDestino Is Nothing Or Not ActiveSheet Is wsDestino Then
        MsgBox "Esta macro solo puede ejecutarse desde la hoja 'RESIDENCIA OVIEDO'.", vbCritical
        Exit Sub
    End If

    ' --- 2) Selecci�n archivo origen ---
    rutaOrigen = Application.GetOpenFilename("Archivos Excel (*.xls*), *.xls*", , "Seleccione el archivo ORIGEN de datos")
    If rutaOrigen = False Then Exit Sub

    ' --- 3) Desactivar eventos y c�lculo para optimizaci�n ---
    On Error Resume Next
    eventosPrevios = Application.enableEvents: Application.enableEvents = False
    pantallaPrevio = Application.screenUpdating: Application.screenUpdating = False
    calcPrevio = Application.calculation: Application.calculation = xlCalculationManual
    On Error GoTo 0

    ' --- 4) Abrir libro origen si no est� abierto ---
    abiertoPorMacro = False
    On Error Resume Next
    Set wbOrigen = Workbooks(Dir(rutaOrigen))
    On Error GoTo 0
    If wbOrigen Is Nothing Then
        Set wbOrigen = Workbooks.Open(Filename:=rutaOrigen, ReadOnly:=True)
        abiertoPorMacro = True
    End If

    ' --- 5) Validar hoja origen ---
    On Error Resume Next
    Set wsOrigen = wbOrigen.Worksheets("REG-UNIDAD 2025")
    On Error GoTo 0
    If wsOrigen Is Nothing Then
        MsgBox "No se ha encontrado la hoja 'REG-UNIDAD 2025' en el archivo ORIGEN.", vbCritical
        GoTo LimpiezaFinal
    End If

    ' --- 6) Determinar �ltima fila con datos en columna B ---
    ultimaFilaOrigen = LastRowInColumn(wsOrigen, "B", 5)
    If ultimaFilaOrigen < 5 Then
        MsgBox "No hay datos suficientes en la hoja ORIGEN (desde fila 5).", vbExclamation
        GoTo LimpiezaFinal
    End If
    numFilas = ultimaFilaOrigen - 5 + 1

    ' --- 7) Definir mapeo columnas ORIGEN -> DESTINO ---
    columnasMapeo = Array("A", "X", "U", "J", "E", "D", "F", "K", "L", "V", "H", "AB", "AE", "AF", "AG", "I", "AD", "W", "M", "N", "Y", "Z", "S", "T")
    columnasDestino = Array("A", "B", "C", "G", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AC")
    ' usar �ndices expl�citos para ReDim (evita problemas ByRef/Variant)
    Dim lowIdx As Long, highIdx As Long
    lowIdx = LBound(columnasDestino)
    highIdx = UBound(columnasDestino)
    ReDim columnasFormulas(lowIdx To highIdx)

    ' --- 8) Identificar columnas con f�rmula en fila 2 ---
    For idx = lowIdx To highIdx
        columnasFormulas(idx) = wsDestino.Cells(2, ColLetterToNum(CStr(columnasDestino(idx)))).HasFormula
    Next idx

    Application.StatusBar = "Copiando datos, por favor espere..."

    ' --- 9) Copiar valores desde origen a destino (con tratamiento de fechas)
    ' - Origen filas: 5 .. ultimaFilaOrigen
    ' - Destino filas: 2 .. 1 + numFilas
    For idx = lowIdx To highIdx
        ' Convertir letras a n�meros de columna (siempre pasar CStr para evitar ByRef issues)
        colONum = ColLetterToNum(CStr(columnasMapeo(idx)))    ' columna en ORIGEN (n�mero)
        colDNum = ColLetterToNum(CStr(columnasDestino(idx)))  ' columna en DESTINO (n�mero)

        If Not columnasFormulas(idx) Then
            With wsDestino.Range(wsDestino.Cells(2, colDNum), wsDestino.Cells(1 + numFilas, colDNum))
                ' Tratamiento especial para fechas en columnas L, M y B (B viene de X)
                If CStr(columnasDestino(idx)) = "L" Or CStr(columnasDestino(idx)) = "M" Or CStr(columnasDestino(idx)) = "B" Then
                    arrFechas = wsOrigen.Range(wsOrigen.Cells(5, colONum), wsOrigen.Cells(ultimaFilaOrigen, colONum)).Value
                    ' Normalizar valores tipo fecha a Date en memoria
                    If IsArray(arrFechas) Then
                        For r = 1 To UBound(arrFechas, 1)
                            If IsDate(arrFechas(r, 1)) Then
                                arrFechas(r, 1) = CDate(arrFechas(r, 1))
                            End If
                        Next r
                    Else
                        ' Si solo una celda (1 fila), arrFechas no es array; tratarlo como valor �nico
                        If IsDate(arrFechas) Then
                            arrFechas = CDate(arrFechas)
                        End If
                    End If
                    .Value = arrFechas
                    ' Formato: B incluye horas, L y M solo fecha
                    If CStr(columnasDestino(idx)) = "B" Then
                        .NumberFormat = "dd/mm/yyyy hh:mm"
                    Else
                        .NumberFormat = "dd/mm/yyyy"
                    End If
                Else
                    ' Copia directa para resto de columnas (r�pido y en bloque)
                    .Value = wsOrigen.Range(wsOrigen.Cells(5, colONum), wsOrigen.Cells(ultimaFilaOrigen, colONum)).Value
                End If
            End With
        End If
    Next idx

    ' --- 10) Limpiar columna T seg�n valores de resoluci�n en P ---
    Dim rngT As Range, rngP As Range
    Set rngT = wsDestino.Range("T2:T" & 1 + numFilas)
    Set rngP = wsDestino.Range("P2:P" & 1 + numFilas)
    For i = 1 To numFilas
        If UCase(Trim(rngP.Cells(i, 1).Value)) = "NO" Or _
           UCase(Trim(rngP.Cells(i, 1).Value)) = "RENUNCIA" Or _
           UCase(Trim(rngP.Cells(i, 1).Value)) = "DESESTIMADA" Or _
           UCase(Trim(rngP.Cells(i, 1).Value)) = "DENEGADA" Then
            rngT.Cells(i, 1).ClearContents
        End If
    Next i

    ' --- 11) Reglas especiales G/D/E/F ---
    For i = 1 To numFilas
        Dim valorG As String, valorGKey As String
        valorG = CStr(wsOrigen.Cells(i + 4, "J").Value) ' Ajustado a columna J en origen (seg�n mapeo)
        valorGKey = Replace(Trim(valorG), " ", "")
        If dictG.Exists(valorGKey) Then
            posG = dictG(valorGKey)
            If Not wsDestino.Cells(2, ColLetterToNum("D")).HasFormula Then wsDestino.Cells(i + 1, ColLetterToNum("D")).Value = listaD(posG)
            If Not wsDestino.Cells(2, ColLetterToNum("E")).HasFormula Then wsDestino.Cells(i + 1, ColLetterToNum("E")).Value = listaE(posG)
            If Not wsDestino.Cells(2, ColLetterToNum("F")).HasFormula Then wsDestino.Cells(i + 1, ColLetterToNum("F")).Value = listaF(posG)
        End If
    Next i

    ' --- 12) Recalcular columna H seg�n reglas de Finalidad ---
    Call RecalcularColumnaH(wsDestino, 2, 1 + numFilas)

    ' --- 13) Replicar autom�ticamente A ? AC ---
    Dim rngA As Range, celda As Range
    Set rngA = wsDestino.Range("A2:A" & 1 + numFilas)
    For Each celda In rngA
        If IsEmpty(celda.Value) Then
            wsDestino.Cells(celda.Row, ColLetterToNum("AC")).ClearContents
        Else
            wsDestino.Cells(celda.Row, ColLetterToNum("AC")).Value = celda.Value
        End If
    Next celda

    ' --- 14) Columna AB = "SI" si columna C tiene valor num�rico ---
    Dim rngC As Range
    Set rngC = wsDestino.Range("C2:C" & 1 + numFilas)
    For Each celda In rngC
        If IsNumeric(celda.Value) And Not IsEmpty(celda.Value) Then
            wsDestino.Cells(celda.Row, ColLetterToNum("AB")).Value = "SI"
        End If
    Next celda

    Application.StatusBar = False
    MsgBox "La copia de datos se realiz� con �xito.", vbInformation

LimpiezaFinal:
    On Error Resume Next
    If abiertoPorMacro And Not wbOrigen Is Nothing Then wbOrigen.Close SaveChanges:=False
    Application.enableEvents = eventosPrevios
    Application.screenUpdating = pantallaPrevio
    Application.calculation = calcPrevio
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' === Recalcular columna H (COMISI�N) a partir de D ===
Private Sub RecalcularColumnaH(ws As Worksheet, filaInicio As Long, filaFin As Long)
    Dim i As Long, valorFinalidad As String
    For i = filaInicio To filaFin
        valorFinalidad = UCase(Trim(ws.Cells(i, ColLetterToNum("D")).Value))
        Select Case valorFinalidad
            Case "COMISI�N", "COMISI�N NO INDEM.", "DESTINO"
                ws.Cells(i, ColLetterToNum("H")).Value = "SI"
            Case ""
                ws.Cells(i, ColLetterToNum("H")).ClearContents
            Case Else
                ws.Cells(i, ColLetterToNum("H")).Value = "NO"
        End Select
    Next i
End Sub

' === �ltima fila con valor visible (xlValues) en una columna, desde fila inicial ===
Private Function LastRowInColumn(ws As Worksheet, colLetter As String, Optional startRow As Long = 1) As Long
    Dim rng As Range, f As Range
    Set rng = ws.Range(colLetter & startRow & ":" & colLetter & ws.Rows.Count)
    Set f = rng.Find(What:="*", LookIn:=xlValues, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If Not f Is Nothing Then
        LastRowInColumn = f.Row
    Else
        LastRowInColumn = startRow - 1
    End If
End Function

' === Conversi�n de letra de columna a n�mero ===
Public Function ColLetterToNum(colLetter As String) As Long
    Dim i As Long, result As Long
    colLetter = UCase(Trim(colLetter))
    result = 0
    For i = 1 To Len(colLetter)
        result = result * 26 + (Asc(Mid(colLetter, i, 1)) - Asc("A") + 1)
    Next i
    ColLetterToNum = result
End Function


