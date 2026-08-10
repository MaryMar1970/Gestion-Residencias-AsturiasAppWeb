Attribute VB_Name = "CalculoDiasUso"

' M�dulo: CalculoDiasUso
Option Explicit

Public Sub CalcularDiasUso(ByVal ws As Worksheet, ByVal Target As Range)

    Dim rngEntrada As Range
    Dim rngSalida As Range
    Dim rngDiasUso As Range
    Dim celda As Range

    ' Definir los rangos de las columnas L (ENTRADA), M (SALIDA) y N (DIAS USO)
    Set rngEntrada = ws.Columns("L")
    Set rngSalida = ws.Columns("M")
    Set rngDiasUso = ws.Columns("N")

    ' Procesar cada celda afectada por el cambio
    For Each celda In Intersect(Target, Union(rngEntrada, rngSalida))
        Dim celdaEntrada As Range
        Dim celdaSalida As Range
        Dim celdaDias As Range

        ' Obtener las celdas correspondientes
        Set celdaEntrada = ws.Cells(celda.Row, rngEntrada.Column)
        Set celdaSalida = ws.Cells(celda.Row, rngSalida.Column)
        Set celdaDias = ws.Cells(celda.Row, rngDiasUso.Column)

        ' Verificar que las fechas son v�lidas
        On Error Resume Next
        If IsDate(celdaEntrada.Value) And IsDate(celdaSalida.Value) Then
            ' Calcular los d�as de uso sin restricciones
            celdaDias.Value = celdaSalida.Value - celdaEntrada.Value
        Else
            celdaDias.Value = ""
        End If
        On Error GoTo 0
    Next celda
End Sub

