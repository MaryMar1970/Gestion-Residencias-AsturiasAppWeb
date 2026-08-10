Attribute VB_Name = "BotonIzdaDchaHojasRESIDENCIAS"
Public Sub DesplazarBloqueHaciaIzquierda(control As IRibbonControl)
    Dim BLOQUE_COLUMNAS As Integer
    Dim nuevaColumna As Long
    ' Intenta leer el tama�o de bloque de la celda B1 de la hoja activa
    If IsNumeric(ActiveSheet.Range("B1").Value) Then
        BLOQUE_COLUMNAS = ActiveSheet.Range("B1").Value
    Else
        BLOQUE_COLUMNAS = 10 ' Valor por defecto
    End If

    Select Case ActiveSheet.Name
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"
            nuevaColumna = ActiveWindow.ScrollColumn - BLOQUE_COLUMNAS
            If nuevaColumna < 1 Then nuevaColumna = 1
            ActiveWindow.ScrollColumn = nuevaColumna
    End Select
End Sub

Public Sub DesplazarBloqueHaciaDerecha(control As IRibbonControl)
    Dim BLOQUE_COLUMNAS As Integer
    Dim maxColumna As Long
    Dim nuevaColumna As Long
    Dim columnasVisible As Long

    ' Intenta leer el tama�o de bloque de la celda B1 de la hoja activa
    If IsNumeric(ActiveSheet.Range("B1").Value) Then
        BLOQUE_COLUMNAS = ActiveSheet.Range("B1").Value
    Else
        BLOQUE_COLUMNAS = 10
    End If

    Select Case ActiveSheet.Name
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"
            maxColumna = ActiveSheet.Columns.Count
            columnasVisible = ActiveWindow.VisibleRange.Columns.Count
            nuevaColumna = ActiveWindow.ScrollColumn + BLOQUE_COLUMNAS
            If nuevaColumna > (maxColumna - columnasVisible + 1) Then
                nuevaColumna = maxColumna - columnasVisible + 1
                If nuevaColumna < 1 Then nuevaColumna = 1
            End If
            ActiveWindow.ScrollColumn = nuevaColumna
    End Select
End Sub
