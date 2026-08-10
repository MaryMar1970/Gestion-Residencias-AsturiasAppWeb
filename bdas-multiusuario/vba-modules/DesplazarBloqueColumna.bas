Attribute VB_Name = "DesplazarBloqueColumna"
' ==== M�dulo: DesplazarBloqueHastaColumnaX ====
' Permite desplazarse horizontalmente en bloques por las hojas de residencias

Option Explicit

Public Sub DesplazarColumnasHaciaDerecha(control As Object)
    Const BLOQUE_COLUMNAS As Long = 17
    Dim maxColumna As Long
    Dim nuevaColumna As Long
    Dim columnasVisible As Long

    If ActiveWindow Is Nothing Then Exit Sub

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


Public Sub DesplazarColumnasHaciaIzquierda(control As Object)
    Const BLOQUE_COLUMNAS As Long = 17
    Dim nuevaColumna As Long

    If ActiveWindow Is Nothing Then Exit Sub

    Select Case ActiveSheet.Name
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"
            nuevaColumna = ActiveWindow.ScrollColumn - BLOQUE_COLUMNAS
            If nuevaColumna < 1 Then nuevaColumna = 1
            ActiveWindow.ScrollColumn = nuevaColumna
    End Select
End Sub

