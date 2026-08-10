Attribute VB_Name = "BorradoNumOrden"
' =============================================================================
' M�DULO: BorradoNumOrden
' Centraliza la confirmaci�n y ejecuci�n del borrado de fila cuando se elimina
' el N� ORDEN (columna A) en las hojas RESIDENCIA GIJ�N, SOTO y OVIEDO.
'
' LLAMADA DESDE: Case 1 del Worksheet_Change de cada hoja de residencia,
'                �nicamente cuando Target est� vac�o tras el borrado.
'
' PAR�METROS:
'   ws               - La hoja desde la que se llama (Me)
'   fila             - Fila donde se ha borrado el N� ORDEN
'   numOrdenAnterior - Valor previo de la celda A (ValorAnteriorXxx del libro)
'   nombreResumen    - Nombre de la hoja RESUMEN correspondiente
'   ultimaColumna    - Letra de la �ltima columna de datos de esa hoja ("AG", "AB"...)
' =============================================================================
Option Explicit

Public Sub GestionarBorradoOrden( _
    ByVal ws As Worksheet, _
    ByVal fila As Long, _
    ByVal numOrdenAnterior As Variant, _
    ByVal nombreResumen As String, _
    ByVal ultimaColumna As String)

    Dim numOrdenStr As String
    numOrdenStr = Trim(CStr(numOrdenAnterior))

    ' --- Mostrar confirmaci�n ANTES de deshabilitar eventos ---
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox( _
        "Se va a borrar el N� ORDEN: " & numOrdenStr & Chr(13) & Chr(13) & _
        "Esta acci�n eliminar� el contenido de toda la fila." & Chr(13) & _
        "�Desea continuar?", _
        vbYesNo + vbExclamation, _
        "Confirmar borrado de N� ORDEN")

    ' --- Si cancela, deshace el borrado de A y sale ---
    If respuesta = vbNo Then
        Application.enableEvents = False
        ws.Cells(fila, "A").Value = numOrdenAnterior
        Application.enableEvents = True
        ws.Cells(fila, "A").Select
        Exit Sub
    End If

    ' --- Confirma: deshabilita eventos y borra valores C:ultimaColumna ---
    Application.enableEvents = False
    Application.screenUpdating = False

    On Error GoTo ErrorBorrado

    Dim colFin As Long
    colFin = ws.Range(ultimaColumna & "1").Column

    Dim cBorrar As Range
    Dim c As Range
    Set cBorrar = ws.Range(ws.Cells(fila, 2), ws.Cells(fila, colFin))

    For Each c In cBorrar.Cells
        If Not c.HasFormula Then
            If Not IsEmpty(c.Value) Then c.ClearContents
        End If
    Next c

    ' --- Borra la fila correspondiente en la hoja RESUMEN ---
    Call BorrarFilaEnResumen_Central(numOrdenAnterior, nombreResumen)

    GoTo Restaurar

ErrorBorrado:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "BorradoNumOrden"

Restaurar:
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic

End Sub


' =============================================================================
' PRIVADA: BorrarFilaEnResumen_Central
' R�plica interna del BorrarFilaEnResumen de cada hoja, centralizada aqu�
' para no depender de la visibilidad de los m�todos privados de cada hoja.
' =============================================================================
Private Sub BorrarFilaEnResumen_Central( _
    ByVal numOrden As Variant, _
    ByVal nombreHojaResumen As String)

    Dim wsResumen As Worksheet
    Dim celdaEncontrada As Range
    Dim ultimaFila As Long
    Dim numOrdenStr As String

    numOrdenStr = Trim(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    On Error Resume Next
    Set wsResumen = ThisWorkbook.Worksheets(nombreHojaResumen)
    On Error GoTo 0

    If wsResumen Is Nothing Then Exit Sub

    ultimaFila = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then Exit Sub

    On Error Resume Next
    Set celdaEncontrada = wsResumen.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If Not celdaEncontrada Is Nothing Then
        celdaEncontrada.EntireRow.Delete
    End If

End Sub

