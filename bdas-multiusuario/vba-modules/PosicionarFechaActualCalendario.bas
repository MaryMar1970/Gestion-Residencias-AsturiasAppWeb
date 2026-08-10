Attribute VB_Name = "PosicionarFechaActualCalendario"
'=================================================================================
' M�dulo: PosicionarFechaActual
'
' Descripci�n:
'   Proporciona dos procedimientos:
'   1. PosicionarFechaActual: Macro conectada a la cinta de opciones (Ribbon) que
'      posiciona la vista en la columna de la fecha actual, SOLO en hojas de tipo calendario.
'   2. PosicionarFecha: Dado un objeto Worksheet, busca la fecha actual en la fila 1
'      (rango B1:APZ1), y si la encuentra, desplaza la vista a esa celda.
'=================================================================================

Option Explicit

'---------------------------------------------------------------------------------
' PosicionarFechaActual
'   - control: Objeto IRibbonControl (permite integraci�n con la cinta de opciones)
'
'   Si la hoja activa es uno de los calendarios (GIJ�N, SOTO, OVIEDO),
'   llama a PosicionarFecha sobre la hoja activa.
'   Si no, muestra mensaje informativo.
'---------------------------------------------------------------------------------
Private Sub PosicionarFechaActual(control As IRibbonControl)
    ' Verifica si la hoja activa es una de las hojas deseadas
    If ActiveSheet.Name = "Calendario GIJ�N" Or _
       ActiveSheet.Name = "Calendario SOTO" Or _
       ActiveSheet.Name = "Calendario OVIEDO" Then
        ' Llama a la macro principal usando la hoja activa
        PosicionarFecha ActiveSheet
    Else
        MsgBox "Esta macro solo se puede ejecutar en las hojas 'Calendario GIJ�N', 'Calendario SOTO' o 'Calendario OVIEDO'.", vbExclamation
    End If
End Sub

'---------------------------------------------------------------------------------
' PosicionarFecha
'   - ws: Hoja de c�lculo en la que buscar la fecha actual
'
'   Busca la fecha actual en la fila 1 (rango B1:APZ1) de la hoja 'ws'.
'   Si la encuentra, desplaza la vista a esa celda. Si no, muestra aviso.
'---------------------------------------------------------------------------------
Public Sub PosicionarFecha(ws As Worksheet)
    Dim currentDate As Date         ' Fecha actual del sistema
    Dim cell As Range               ' Celda en el bucle de b�squeda
    Dim found As Boolean            ' Marca si se encontr� la fecha

    currentDate = Date              ' Obtiene la fecha del sistema (sin hora)

    found = False
    ' Recorre las celdas de la fila 1 (de B1 a APZ1)
    For Each cell In ws.Range("B1:APZ1")
        If cell.Value = currentDate Then
            ' Si encuentra la fecha, desplaza la vista a esa celda
            Application.GoTo Reference:=cell, Scroll:=True
            found = True
            Exit For
        End If
    Next cell

    ' Si no se encontr� la fecha en la fila 1, muestra mensaje informativo
    If Not found Then
        MsgBox "La fecha actual no se encuentra en el cuadrante.", vbInformation
    End If
End Sub

