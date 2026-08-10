Attribute VB_Name = "FacturacionMesSOTO"
'=================================================================================
' Macro principal: ActualizarFacturacionFacturaSoto
'---------------------------------------------------------------------------------
' Actualiza FACTURACI�N SOTO seg�n datos de RESIDENCIA SOTO y mes de B2.
' Filtra, valida, convierte importes, ordena, calcula acumulados, y maneja errores.
' Opciones y estructura documentadas para mantenimiento futuro.
'=================================================================================
Option Explicit

Sub ActualizarFacturacionFacturaSoto()
    Dim wsRes As Worksheet, wsFac As Worksheet
    Dim lastRowRes As Long, lastRowFacClear As Long
    Dim selectedMonthName As String, selectedMonthNum As Integer
    Dim dataRes As Variant, outputData() As Variant, finalOutput() As Variant
    Dim matchCount As Long, i As Long, currentIndex As Long
    Dim cellDate As Variant, cellResolucion As String, cellPagado As String, defaultFontSize As Single
    Dim avisoCampos As String

    ' Desactiva actualizaciones y eventos para m�ximo rendimiento y evitar recursividad
    ' Desactiva actualizaciones y eventos para mximo rendimiento y evitar recursividad
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False

    On Error GoTo CleanUp

    ' Referencia hojas y sincronizar desde Access DB
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA SOTO")
    Set wsFac = ThisWorkbook.Worksheets("FACTURACIN SOTO")
    
    modDatabase.SincronizarHojaDesdeAccess "SOTO", wsRes

    ' Lee y valida el mes seleccionado en B2
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

    ' Limpia registros previos (A4:H...)
    lastRowFacClear = wsFac.Cells(wsFac.Rows.Count, "A").End(xlUp).Row
    If lastRowFacClear >= 4 Then
        wsFac.Range("A4:I" & lastRowFacClear).ClearContents
    End If

    ' Lee todo el rango de datos de RESIDENCIA SOTO a memoria (array)
    lastRowRes = wsRes.Cells(wsRes.Rows.Count, "M").End(xlUp).Row
    If lastRowRes < 2 Then GoTo CleanUp
    dataRes = wsRes.Range("A2:AA" & lastRowRes).Value

    ' Cuenta coincidencias v�lidas (filtrado por mes, resoluci�n, pagado)
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

    ' Si no hay registros v�lidos, avisa y termina
    If matchCount = 0 Then
        If Application.ActiveSheet.Name = wsFac.Name Then
            MsgBox "No se encontraron datos para el mes seleccionado.", vbInformation
        End If
        GoTo CleanUp
    End If

    ' Prepara salida filtrada (7 columnas)
    ReDim outputData(1 To matchCount, 1 To 8)
    currentIndex = 1
    avisoCampos = ""

    ' Segunda pasada: filtra y valida campos clave
    For i = 1 To UBound(dataRes, 1)
        If IsDate(dataRes(i, 13)) Then
            cellDate = CDate(dataRes(i, 13))
            If Month(cellDate) = selectedMonthNum Then
                cellResolucion = UCase(Trim(CStr(dataRes(i, 16))))
                cellPagado = UCase(Trim(CStr(dataRes(i, 27))))
                If (cellResolucion = "SI" Or cellResolucion = "CONCEDIDA" Or cellResolucion = "REEVALUADA") And MarcarSiPagadosEnResidencia.EsEstadoPagado(cellPagado) Then
                    If IsEmpty(dataRes(i, 1)) Then avisoCampos = avisoCampos & "Falta N� ORDEN en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 3)) Then avisoCampos = avisoCampos & "Falta N� FACTURA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 9)) Then avisoCampos = avisoCampos & "Falta DNI en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 11)) Then avisoCampos = avisoCampos & "Falta NOMBRE en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 12)) Then avisoCampos = avisoCampos & "Falta FECHA ENTRADA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 13)) Then avisoCampos = avisoCampos & "Falta FECHA SALIDA en fila " & i + 1 & vbNewLine
                    If IsEmpty(dataRes(i, 21)) Then avisoCampos = avisoCampos & "Falta IMPORTE en fila " & i + 1 & vbNewLine

                    ' Conversi?n robusta de IMPORTE
                    Dim importeVal As Double
                    If IsNumeric(dataRes(i, 21)) Then
                        importeVal = CDbl(dataRes(i, 21))
                    ElseIf Len(dataRes(i, 21)) > 0 Then
                        importeVal = val(Replace(Replace(dataRes(i, 21), "?", ""), ",", "."))
                    Else
                        importeVal = 0
                    End If

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

    ' Avisos agrupados si faltan campos
    If Len(avisoCampos) > 0 Then
        MsgBox "Avisos sobre datos faltantes:" & vbNewLine & avisoCampos, vbExclamation
    End If

    ' Ordena por fecha de salida usando hoja temporal
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

    ' Calcula acumulados y prepara array final
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

    ' Vuelca el resultado final en FACTURACI�N SOTO desde A4
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

