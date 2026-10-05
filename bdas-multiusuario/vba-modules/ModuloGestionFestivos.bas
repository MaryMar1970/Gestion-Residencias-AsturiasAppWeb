Attribute VB_Name = "ModuloGestionFestivos"
Option Explicit

' ==============================================================================
' ModuloGestionFestivos: Gestión de días festivos diferenciados por residencia
' ==============================================================================
' Estructura de la hoja 'Festivos':
'   - Columna A: FESTIVOS GIJÓN
'   - Columna B: FESTIVOS SOTO
'   - Columna C: FESTIVOS OVIEDO
' ==============================================================================

''' Colorea únicamente las fechas festivas correspondientes a cada residencia en:
''' "Calendario GIJÓN", "Calendario SOTO" y "Calendario OVIEDO".
''' Si un festivo se retira de su columna, se limpia automáticamente su color.
Public Sub ResaltarFestivos(Optional ByVal wsTarget As Worksheet = Nothing)
    On Error Resume Next
    Dim wsFest As Worksheet
    Set wsFest = ThisWorkbook.Sheets("Festivos")
    If wsFest Is Nothing Then Exit Sub
    
    Dim colorFestivo As Long
    colorFestivo = RGB(255, 120, 120)  ' Rojo suave
    
    Dim wsList As Collection
    Set wsList = New Collection
    
    If Not wsTarget Is Nothing Then
        wsList.Add wsTarget
    Else
        Dim wsItem As Worksheet
        For Each wsItem In ThisWorkbook.Worksheets
            wsList.Add wsItem
        Next wsItem
    End If
    
    Dim ws As Worksheet, wIdx As Long
    For wIdx = 1 To wsList.Count
        Set ws = wsList(wIdx)
        Dim filaExtra As Long
        Dim colFestivos As String
        filaExtra = 0
        colFestivos = ""
        
        ' Identificar calendario y asignar su fila de separación y columna de festivos
        If (InStr(UCase(ws.Name), "GIJÓN") > 0 Or InStr(UCase(ws.Name), "GIJON") > 0) And InStr(UCase(ws.Name), "CALENDARIO") > 0 Then
            filaExtra = 27
            colFestivos = "A"   ' Festivos GIJÓN
        ElseIf InStr(UCase(ws.Name), "SOTO") > 0 And InStr(UCase(ws.Name), "CALENDARIO") > 0 Then
            filaExtra = 20
            colFestivos = "B"   ' Festivos SOTO
        ElseIf InStr(UCase(ws.Name), "OVIEDO") > 0 And InStr(UCase(ws.Name), "CALENDARIO") > 0 Then
            filaExtra = 31
            colFestivos = "C"   ' Festivos OVIEDO
        End If
        
        If filaExtra > 0 And colFestivos <> "" Then
            ' 1. Cargar en memoria los festivos específicos de esta residencia
            Dim dictFestivos As Object
            Set dictFestivos = CreateObject("Scripting.Dictionary")
            Dim ultFilaFest As Long, r As Long
            ultFilaFest = wsFest.Cells(wsFest.Rows.Count, colFestivos).End(xlUp).Row
            
            If ultFilaFest >= 2 Then
                Dim arrFest As Variant
                arrFest = wsFest.Range(wsFest.Cells(2, colFestivos), wsFest.Cells(ultFilaFest, colFestivos)).Value
                If IsArray(arrFest) Then
                    For r = 1 To UBound(arrFest, 1)
                        If IsDate(arrFest(r, 1)) Then
                            Dim fKey As Long
                            fKey = CLng(CDate(arrFest(r, 1)))
                            If Not dictFestivos.Exists(fKey) Then dictFestivos.Add fKey, True
                        End If
                    Next r
                ElseIf IsDate(arrFest) Then
                    dictFestivos.Add CLng(CDate(arrFest)), True
                End If
            End If
            
            ' 2. Procesar las columnas de fechas del calendario (bloques de 3 columnas)
            Dim ultimaCol As Long, col As Long
            ultimaCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
            If ultimaCol >= 2 Then
                Dim arrRow1 As Variant
                arrRow1 = ws.Range(ws.Cells(1, 1), ws.Cells(1, ultimaCol)).Value
                
                For col = 2 To ultimaCol Step 3
                    If col <= UBound(arrRow1, 2) Then
                        If IsDate(arrRow1(1, col)) Then
                            Dim fCol As Long
                            fCol = CLng(CDate(arrRow1(1, col)))
                            If dictFestivos.Exists(fCol) Then
                                ' Pintar celda de fecha y fila de separación
                                ws.Cells(1, col).Resize(1, 3).Interior.color = colorFestivo
                                ws.Cells(filaExtra, col).Resize(1, 3).Interior.color = colorFestivo
                            Else
                                ' Si ya no es festivo y estaba marcado de rojo, limpiarlo a fondo blanco
                                If ws.Cells(1, col).Interior.color = colorFestivo Then
                                    ws.Cells(1, col).Resize(1, 3).Interior.ColorIndex = xlNone
                                End If
                                If ws.Cells(filaExtra, col).Interior.color = colorFestivo Then
                                    ws.Cells(filaExtra, col).Resize(1, 3).Interior.ColorIndex = xlNone
                                End If
                            End If
                        End If
                    End If
                Next col
            End If
        End If
    Next wIdx
End Sub

''' Devuelve True si la fecha es festivo para la residencia indicada según la hoja 'Festivos'
Public Function EsFestivo(fecha As Date, Optional ByVal residencia As String = "GIJON") As Boolean
    Dim ws As Worksheet, colFestivos As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Festivos")
    If ws Is Nothing Then EsFestivo = False: Exit Function
    
    Dim resNorm As String
    resNorm = UCase(Trim(residencia))
    If InStr(resNorm, "GIJ") > 0 Then
        colFestivos = "A"
    ElseIf InStr(resNorm, "SOTO") > 0 Then
        colFestivos = "B"
    ElseIf InStr(resNorm, "OVIEDO") > 0 Then
        colFestivos = "C"
    Else
        colFestivos = "A"
    End If
    
    Dim lastR As Long
    lastR = ws.Cells(ws.Rows.Count, colFestivos).End(xlUp).Row
    If lastR < 2 Then EsFestivo = False: Exit Function
    
    Dim arr As Variant, i As Long
    arr = ws.Range(ws.Cells(2, colFestivos), ws.Cells(lastR, colFestivos)).Value
    If Not IsArray(arr) Then
        If IsDate(arr) Then EsFestivo = (CLng(CDate(arr)) = CLng(fecha))
        Exit Function
    End If
    
    For i = 1 To UBound(arr, 1)
        If IsDate(arr(i, 1)) Then
            If CLng(CDate(arr(i, 1))) = CLng(fecha) Then
                EsFestivo = True
                Exit Function
            End If
        End If
    Next i
    EsFestivo = False
End Function
