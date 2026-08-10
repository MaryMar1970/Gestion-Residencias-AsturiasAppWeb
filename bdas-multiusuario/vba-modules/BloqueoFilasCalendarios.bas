Attribute VB_Name = "BloqueoFilasCalendarios"
Option Explicit

'-----------------------------------------------------------------------------------
' ProtegerYFormatearFilasCalendarios
'
' Versi�n sencilla y robusta basada en tu c�digo inicial, ampliada a filas adicionales.
' Protege la fila 1 y, adem�s:
'     - GIJ�N: fila 21 (B21:APZ21)
'     - SOTO:  fila 20 (B20:APZ20)
'     - OVIEDO: fila 31 (B31:APZ31)
' No requiere contrase�a.
'-----------------------------------------------------------------------------------
Sub ProtegerYFormatearFilasCalendarios()
    Dim nombresHojas As Variant
    Dim ws As Worksheet
    Dim clave As String
    Dim i As Integer, col As Integer
    Dim inicioBloque As Integer, finBloque As Integer
    Dim mergeArea As Range
    Dim filaActual As Long
    ' Nombres de hojas y filas a proteger
    nombresHojas = Array("Calendario GIJ�N", "Calendario SOTO", "Calendario OVIEDO")
    clave = ""
    Application.screenUpdating = False
    Application.enableEvents = False
    For i = LBound(nombresHojas) To UBound(nombresHojas)
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(nombresHojas(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
        If ws.visible = xlSheetVisible Then
            Dim filasProcesar As Variant
            Select Case ws.Name
                Case "Calendario GIJ�N": filasProcesar = Array(1, 21)
                Case "Calendario SOTO": filasProcesar = Array(1, 20)
                Case "Calendario OVIEDO": filasProcesar = Array(1, 31)
                Case Else: filasProcesar = Array(1)
            End Select
            If ws.ProtectContents Then ws.Unprotect password:=clave
            ws.Cells.Locked = False
            inicioBloque = ws.Range("E1").Column
            finBloque = ws.Range("AQB1").Column
            Dim f As Variant
            For Each f In filasProcesar
                filaActual = f
                ' Bloquea el rango de fecha E:AQB en una sola operacion instantanea (OPTIMIZADO)
                ws.Range(ws.Cells(filaActual, inicioBloque), ws.Cells(filaActual, finBloque)).Locked = True
                ' Formato y protecci�n adecuada de fecha en B:D de la fila actual
                With ws.Range(ws.Cells(filaActual, "B"), ws.Cells(filaActual, "D"))
                    .NumberFormat = "dd/mm/yyyy"
                    If filaActual = 1 Then
                        .Locked = False   ' Solo en la fila 1 se desbloquea la fecha
                    Else
                        .Locked = True    ' En las filas 21, 20, 31 se protege la fecha
                    End If
                End With
            Next f
            ws.Protect password:=clave, UserInterfaceOnly:=True, AllowFormattingCells:=True
        End If
        End If
    Next i
    Application.enableEvents = True
    Application.screenUpdating = True
End Sub
