Attribute VB_Name = "BotonIrUltimaCeldaOrden"
Option Explicit

Sub IrUltimaCeldaVacia(control As IRibbonControl)
    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Select Case ws.Name
        ' ====== Hojas RESIDENCIA: Selecciona PRIMERA CELDA VAC�A de la columna A ======
        Case "RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"
            Dim celdaA As Range
            Set celdaA = ws.Range("A:A").Find(What:="*", After:=ws.Range("A1"), LookIn:=xlValues, _
                                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
            If Not celdaA Is Nothing Then
                If celdaA.Row + 1 <= 4000 Then   ' <--- Cambia 4000 por la �ltima fila �til, si corresponde
                    ws.Cells(celdaA.Row + 1, 1).Select
                Else
                    MsgBox "No hay m�s filas disponibles para seleccionar.", vbExclamation, "Fin de hoja"
                End If
            Else
                ws.Cells(2, 1).Select ' Empieza normalmente en A2, suponiendo encabezado en A1
            End If

        ' ====== Hojas RESUMEN: Selecciona la �LTIMA OCUPADA VISIBLE de la columna B ======
        Case "RESUMEN GIJ�N", "RESUMEN SOTO", "RESUMEN OVIEDO"
            Dim UltFila As Long
            Dim c As Range
            For UltFila = 4000 To 2 Step -1    ' Cambia 4000 si el tama�o cambia
                Set c = ws.Cells(UltFila, "B")
                If Not IsError(c.Value) Then
                    If Len(Trim(c.Value)) > 0 Then
                        c.Select
                        Exit Sub
                    End If
                End If
            Next UltFila
            ' Si no se encuentra nada, posiciona en B2
            ws.Cells(2, "B").Select

        ' ====== Otras hojas: Nada ======
        Case Else
            Exit Sub
    End Select

    Exit Sub

ErrHandler:
    MsgBox "Se produjo un error inesperado: " & Err.Description, vbCritical, "Error inesperado"
End Sub
