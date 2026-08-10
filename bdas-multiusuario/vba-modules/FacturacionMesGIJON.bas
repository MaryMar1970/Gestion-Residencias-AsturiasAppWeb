Attribute VB_Name = "FacturacionMesGIJON"
'=================================================================================
' Macro principal: ActualizarFacturacionFacturaGijon
'---------------------------------------------------------------------------------
' Descripci?n:
'   - Actualiza la hoja FACTURACI�N GIJ�N seg�n los datos de RESIDENCIA GIJ�N y el mes de B2.
'   - Filtra, valida campos clave, convierte importes, ordena y calcula acumulados.
'   - Incluye manejo robusto de errores y restauraci?n del estado de Excel.
'   - Muestra avisos agrupados si faltan campos clave en los registros v�lidos.
'   - Realiza el volcado final en A4:H (m�s columna de acumulado).
'
' Mantenimiento:
'   - Si cambias las columnas en tus hojas, revisa los ?ndices en los arrays.
'   - Si cambian los criterios de filtro, ajusta la l�gica de filtrado.
'   - Si modificas la estructura, actualiza la documentaci?n.
'=================================================================================
Option Explicit

Sub ActualizarFacturacionFacturaGijon()
    Dim wsRes As Worksheet, wsFac As Worksheet
    Dim lastRowRes As Long, lastRowFacClear As Long
    Dim selectedMonthName As String, selectedMonthNum As Integer
    Dim dataRes As Variant, outputData() As Variant, finalOutput() As Variant
    Dim matchCount As Long, i As Long, currentIndex As Long
    Dim cellDate As Variant, cellResolucion As String, cellPagado As String, defaultFontSize As Single
    Dim avisoCampos As String

    '-------------------------------------------------------------
    ' Optimizaci�n: desactiva actualizaciones y eventos para mayor rendimiento
    '-------------------------------------------------------------
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False

    On Error GoTo CleanUp

    '-------------------------------------------------------------
    ' Referenciar hojas de trabajo y sincronizar desde Access DB
    '-------------------------------------------------------------
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA GIJN")
    Set wsFac = ThisWorkbook.Worksheets("FACTURACIN GIJN")
    
    modDatabase.SincronizarHojaDesdeAccess "GIJON", wsRes

    '-------------------------------------------------------------
    ' Leer el mes seleccionado en B2 y validarlo
    '-------------------------------------------------------------
    selectedMonthName = Trim(wsFac.Range("B2").Value)
    Select Case UCase(selectedMonthName)
        Case "ENERO":      selectedMonthNum = 1
        Case "FEBRERO":    selectedMonthNum = 2
        Case "MARZO":      selectedMonthNum = 3
        Case "ABRIL":      selectedMonthNum = 4
        Case "MAYO":       selectedMonthNum = 5
        Case "JUNIO":      selectedMonthNum = 6
        Case "JULIO":      selectedMonthNum = 7
        Case "AGOSTO":     selectedMonthNum = 8
        Case "SEPTIEMBRE": selectedMonthNum = 9
        Case "OCTUBRE":    selectedMonthNum = 10
        Case "NOVIEMBRE":  selectedMonthNum = 11
        Case "DICIEMBRE":  selectedMonthNum = 12
        Case Else
            MsgBox "El mes seleccionado en B2 no es v�lido.", vbExclamation
            GoTo CleanUp
    End Select

    '-------------------------------------------------------------
    ' Limpiar registros previos en FACTURACI�N GIJ�N (A4:H...)
    '-------------------------------------------------------------
    lastRowFacClear = wsFac.Cells(wsFac.Rows.Count, "A").End(xlUp).Row
    If lastRowFacClear >= 4 Then
        wsFac.Range("A4:I" & lastRowFacClear).ClearContents
    End If

    '-------------------------------------------------------------
    ' Leer todo el rango de datos de RESIDENCIA GIJ�N a memoria (array)
    ' Nota: Lee hasta columna AA (col 27)
    '-------------------------------------------------------------
    lastRowRes = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRowRes < 2 Then GoTo CleanUp

    dataRes = wsRes.Range("A2:AA" & lastRowRes).Value

    '-------------------------------------------------------------
    ' Primera pasada: contar cu?ntos registros cumplen los criterios del mes y estado
    '-------------------------------------------------------------
    matchCount = 0
    For i = 1 To UBound(dataRes, 1)
        If IsDate(dataRes(i, 13)) Then
            cellDate = CDate(dataRes(i, 13))
            If Month(cellDate) = selectedMonthNum Then
                cellResolucion = UCase(Trim(CStr(dataRes(i, 16))))
                cellPagado = UCase(Trim(CStr(dataRes(i, 27))))
                If (cellResolucion = "SI" Or cellResolucion = "CONCEDIDA" Or cellResolucion = "REEVALUADA") And MarcarSiPagadosEnResidencia.EsEstadoPagado(cellPagado) Then
                    matchCount = matchCount + 1
                End If
            End If
        End If
    Next i

    '-------------------------------------------------------------
    ' Si no hay datos v�lidos, avisa y termina
    '-------------------------------------------------------------
    If matchCount = 0 Then
        If Application.ActiveSheet.Name = wsFac.Name Then
            MsgBox "No se encontraron datos para el mes seleccionado.", vbInformation
        End If
        GoTo CleanUp
    End If

    '-------------------------------------------------------------
    ' Redimensionar el array de salida con el n�mero exacto de coincidencias (7 columnas)
    '-------------------------------------------------------------
    ReDim outputData(1 To matchCount, 1 To 8)
    currentIndex = 1
    avisoCampos = ""

    '-------------------------------------------------------------
    ' Segunda pasada: llenar el array con los datos que cumplen el criterio
    ' Validar y agrupar avisos por campos clave faltantes
    '-------------------------------------------------------------
    For i = 1 To UBound(dataRes, 1)
        If IsDate(dataRes(i, 13)) Then
            cellDate = CDate(dataRes(i, 13))
            If Month(cellDate) = selectedMonthNum Then
                cellResolucion = UCase(Trim(CStr(dataRes(i, 16))))
                cellPagado = UCase(Trim(CStr(dataRes(i, 27))))
                If (cellResolucion = "SI" Or cellResolucion = "CONCEDIDA" Or cellResolucion = "REEVALUADA") And MarcarSiPagadosEnResidencia.EsEstadoPagado(cellPagado) Then

                    ' Validaci�n de campos clave
                    If IsEmpty(dataRes(i, 1)) Then avisoCampos = avisoCampos & "Falta N� ORDEN en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 3)) Then avisoCampos = avisoCampos & "Falta N� FACTURA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 9)) Then avisoCampos = avisoCampos & "Falta DNI en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 11)) Then avisoCampos = avisoCampos & "Falta NOMBRE en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 12)) Then avisoCampos = avisoCampos & "Falta FECHA ENTRADA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 13)) Then avisoCampos = avisoCampos & "Falta FECHA SALIDA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 21)) Then avisoCampos = avisoCampos & "Falta IMPORTE en fila " & i + 1 & vbNewLine

                    ' Conversi?n robusta de IMPORTE: elimina s?mbolo ?, convierte a n�mero
                    Dim importeVal As Double
                    If IsNumeric(dataRes(i, 21)) Then
                        importeVal = CDbl(dataRes(i, 21))
                    ElseIf Len(dataRes(i, 21)) > 0 Then
                        importeVal = val(Replace(Replace(dataRes(i, 21), "?", ""), ",", "."))
                    Else
                        importeVal = 0
                    End If

                    ' Asignar los datos correspondientes:
                    outputData(currentIndex, 1) = dataRes(i, 1)    ' N� ORDEN
                    outputData(currentIndex, 2) = dataRes(i, 3)    ' N� FACTURA
                    outputData(currentIndex, 3) = dataRes(i, 9)    ' DNI
                    outputData(currentIndex, 4) = dataRes(i, 11)   ' NOMBRE
                    outputData(currentIndex, 5) = dataRes(i, 12)   ' FECHA ENTRADA
                    outputData(currentIndex, 6) = dataRes(i, 13)   ' FECHA SALIDA
                    outputData(currentIndex, 7) = importeVal       ' IMPORTE
                    outputData(currentIndex, 8) = dataRes(i, 27)   ' PAGO
                    currentIndex = currentIndex + 1
                End If
            End If
        End If
    Next i

    '-------------------------------------------------------------
    ' Si hay avisos por campos faltantes, mostrar mensaje agrupado
    '-------------------------------------------------------------
    If Len(avisoCampos) > 0 Then
        MsgBox "Avisos sobre datos faltantes:" & vbNewLine & avisoCampos, vbExclamation
    End If

    '-------------------------------------------------------------
    ' Ordenar los registros por fecha de salida usando hoja temporal
    '-------------------------------------------------------------
    Dim tempSheet As Worksheet, tempRange As Range
    On Error Resume Next
    Set tempSheet = ThisWorkbook.Worksheets.Add
    On Error GoTo CleanUp
    With tempSheet
        .Range("A1").Resize(UBound(outputData, 1), UBound(outputData, 2)).Value = outputData
        Set tempRange = .Range("A1").CurrentRegion
        tempRange.Sort Key1:=tempRange.Columns(6), Order1:=xlAscending, Header:=xlNo
        outputData = tempRange.Value
        Application.DisplayAlerts = False
        Application.DisplayAlerts = True
    End With

    '-------------------------------------------------------------
    ' Crear el array final que incluye la suma acumulada en la columna H
    '-------------------------------------------------------------
    ReDim finalOutput(1 To matchCount, 1 To 9)
    Dim acumulado As Double: acumulado = 0
    For i = 1 To matchCount
        finalOutput(i, 1) = outputData(i, 1)
        finalOutput(i, 2) = outputData(i, 2)
        finalOutput(i, 3) = outputData(i, 3)
        finalOutput(i, 4) = outputData(i, 4)
        finalOutput(i, 5) = outputData(i, 5)
        finalOutput(i, 6) = outputData(i, 6)
                finalOutput(i, 7) = outputData(i, 7)
        finalOutput(i, 8) = outputData(i, 8)
        acumulado = acumulado + outputData(i, 7)
        finalOutput(i, 9) = acumulado
    Next i

    '-------------------------------------------------------------
    ' Volcar los datos finales en FACTURACI�N GIJ�N a partir de la celda A4
    '-------------------------------------------------------------
    wsFac.Range("A4").Resize(matchCount, 9).Value = finalOutput

CleanUp:
    On Error Resume Next
    If Not tempSheet Is Nothing Then
        Application.DisplayAlerts = False
        tempSheet.Delete
        Application.DisplayAlerts = True
    End If
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
End Sub

