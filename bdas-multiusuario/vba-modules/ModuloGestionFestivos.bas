Attribute VB_Name = "ModuloGestionFestivos"
Option Explicit

Sub ResaltarFestivos()
    ' ==========================================================================
    ' Procedimiento: ResaltarFestivos
    ' Descripci�n  : Colorea �nicamente las fechas festivas en las hojas
    '                "Calendario GIJ�N", "Calendario SOTO" y "Calendario OVIEDO".
    '                Si un festivo se elimina de la hoja "Festivos", se borra
    '                solo su color (sin afectar otros formatos).
    ' Autor        : Jaime Garc�a + ChatGPT (optimizado)
    ' ==========================================================================
    
    Dim wsFest As Worksheet
    Dim ws As Worksheet
    Dim celda As Range
    Dim fechaFestivo As Date
    Dim dictFestivos As Object
    Dim ultimaFila As Long
    Dim rngFechas As Range, c As Range
    Dim filaExtra As Long
    Dim colorFestivo As Long
    Dim clave As Variant
    
    ' === CONFIGURACI�N ===
    colorFestivo = RGB(255, 120, 120)  ' rojo m�s suave
    
    On Error Resume Next
    Set wsFest = ThisWorkbook.Sheets("Festivos")
    On Error GoTo 0
    If wsFest Is Nothing Then
        MsgBox "No se encontr� la hoja 'Festivos'.", vbExclamation
        Exit Sub
    End If
    
    ' === CREAR LISTA DE FESTIVOS (diccionario) ===
    Set dictFestivos = CreateObject("Scripting.Dictionary")
    ultimaFila = wsFest.Cells(wsFest.Rows.Count, "A").End(xlUp).Row
    
    For Each celda In wsFest.Range("A2:A" & ultimaFila)
        If IsDate(celda.Value) Then
            fechaFestivo = DateValue(celda.Value)
            If Not dictFestivos.Exists(fechaFestivo) Then
                dictFestivos.Add fechaFestivo, True
            End If
        End If
    Next celda
    
    ' === PROCESAR HOJAS DE CALENDARIO ===
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case "Calendario GIJ�N": filaExtra = 27
            Case "Calendario SOTO": filaExtra = 20
            Case "Calendario OVIEDO": filaExtra = 31
            Case Else: GoTo SiguienteHoja
        End Select
        
        ' --- Definir rango de fechas combinadas en la fila 1 ---
        Set rngFechas = ws.Range("B1", ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column))
        
        ' --- Recorrer celdas de la fila 1 ---
        For Each c In rngFechas
            If c.MergeCells Then
                Dim celdaMerge As Range
                Set celdaMerge = c.mergeArea
                Dim valorCelda As Variant
                valorCelda = celdaMerge.Cells(1, 1).Value
                
                If IsDate(valorCelda) Then
                    fechaFestivo = DateValue(valorCelda)
                    
                    ' === Si la fecha est� en Festivos ? pintar ===
                    If dictFestivos.Exists(fechaFestivo) Then
                        celdaMerge.Interior.color = colorFestivo
                        ws.Range(ws.Cells(filaExtra, celdaMerge.Column), _
                                 ws.Cells(filaExtra, celdaMerge.Column + celdaMerge.Columns.Count - 1)).Interior.color = colorFestivo
                    ' === Si NO est� en Festivos ? limpiar color solo de esas celdas ===
                    Else
                        celdaMerge.Interior.ColorIndex = xlNone
                        ws.Range(ws.Cells(filaExtra, celdaMerge.Column), _
                                 ws.Cells(filaExtra, celdaMerge.Column + celdaMerge.Columns.Count - 1)).Interior.ColorIndex = xlNone
                    End If
                End If
            End If
        Next c
        
SiguienteHoja:
    Next ws

   ' MsgBox "Actualizaci�n de festivos completada.", vbInformation
End Sub


'=================================================================================
' Funci�n: EsFestivo
' Descripci�n:
'   - Devuelve True si la fecha es un festivo seg�n la hoja "Festivos"
'=================================================================================
Public Function EsFestivo(fecha As Date) As Boolean
    Dim ws As Worksheet, rng As Range, c As Range
    Set ws = ThisWorkbook.Worksheets("Festivos")
    Set rng = ws.Range("A2:A" & ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)
    
    For Each c In rng
        If IsDate(c.Value) Then
            If CLng(c.Value) = CLng(fecha) Then
                EsFestivo = True
                Exit Function
            End If
        End If
    Next c
    EsFestivo = False
End Function


