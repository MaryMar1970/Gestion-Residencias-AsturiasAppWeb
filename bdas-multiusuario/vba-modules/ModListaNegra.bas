Attribute VB_Name = "ModListaNegra"
'=================================================================================
' M�dulo: ModListaNegra
'
' Descripci�n:
'   Verifica si el valor introducido en la columna I de una hoja de trabajo
'   existe en la hoja "LISTA NEGRA". Si se encuentra, muestra una advertencia
'   incluyendo la fecha de inclusi�n (columna B) y el motivo (columna C).
'=================================================================================

Option Explicit

Public Sub CheckBlacklist(ws As Worksheet, Target As Range)
    Dim wsLista As Worksheet
    Dim rngCambios As Range
    Dim rngLista As Range
    Dim cel As Range, found As Range
    Dim fecha As Variant, motivo As Variant
    Dim lastRow As Long

    ' Verifica existencia de hoja LISTA NEGRA
    On Error Resume Next
    Set wsLista = ThisWorkbook.Sheets("LISTA NEGRA")
    On Error GoTo 0

    If wsLista Is Nothing Then
        MsgBox "Hoja 'LISTA NEGRA' no encontrada.", vbCritical
        Exit Sub
    End If

    ' Filtra cambios solo en la columna I
    Set rngCambios = Intersect(Target, ws.Columns("I"))
    If rngCambios Is Nothing Then Exit Sub

    ' Verifica que hay datos en la hoja LISTA NEGRA
    lastRow = wsLista.Cells(wsLista.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Set rngLista = wsLista.Range("A2:A" & lastRow)

    ' Recorre celdas editadas en columna I
    For Each cel In rngCambios
        If Trim(cel.Value) <> "" Then
            Set found = rngLista.Find(What:=cel.Value, LookIn:=xlValues, LookAt:=xlWhole)
            If Not found Is Nothing Then
                fecha = wsLista.Cells(found.Row, "B").Value
                motivo = wsLista.Cells(found.Row, "C").Value

                MsgBox "El DNI '" & cel.Value & "' est� incluido en la hoja LISTA NEGRA:" & vbCrLf & _
                       "� Fecha de inclusi�n: " & Format(fecha, "dd/mm/yyyy") & vbCrLf & _
                       "� Motivo: " & motivo, vbExclamation, "Advertencia: Registro en Lista Negra"
            End If
        End If
    Next cel
End Sub

