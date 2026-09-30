Attribute VB_Name = "BorradoNumOrden"
' =============================================================================
' MÓDULO: BorradoNumOrden
' Centraliza la confirmación y ejecución del borrado de fila cuando se elimina
' el Nº ORDEN (columna A) en las hojas RESIDENCIA GIJÓN, SOTO y OVIEDO.
'
' LLAMADA DESDE: Case 1 del Worksheet_Change de cada hoja de residencia,
'                únicamente cuando Target está vacío tras el borrado.
'
' PARÁMETROS:
'   ws               - La hoja desde la que se llama (Me)
'   fila             - Fila donde se ha borrado el Nº ORDEN
'   numOrdenAnterior - Valor previo de la celda A (ValorAnteriorXxx del libro)
'   nombreResumen    - Nombre de la hoja RESUMEN correspondiente
'   ultimaColumna    - Letra de la última columna de datos de esa hoja ("AG", "AB"...)
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

    ' --- Mostrar confirmación ANTES de deshabilitar eventos ---
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox( _
        "Se va a borrar el Nº ORDEN: " & numOrdenStr & Chr(13) & Chr(13) & _
        "Esta acción eliminará el contenido de toda la fila." & Chr(13) & _
        "¿Desea continuar?", _
        vbYesNo + vbExclamation, _
        "Confirmar borrado de Nº ORDEN")

    ' --- Si cancela, deshace el borrado de A y sale ---
    If respuesta = vbNo Then
        Application.enableEvents = False
        ws.Cells(fila, "A").Value = numOrdenAnterior
        Application.enableEvents = True
        ws.Cells(fila, "A").Select
        Exit Sub
    End If

    ' --- Confirma: deshabilita eventos y prepara borrado ---
    Application.enableEvents = False
    Application.screenUpdating = False

    On Error GoTo ErrorBorrado

    ' --- SINCRONIZACIÓN ACCESS: Eliminar de la base de datos y registrar en Log ---
    If IsNumeric(numOrdenAnterior) Then
        Dim claveRes As String
        claveRes = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
        Call modDatabase.EliminarOrdenBD(CLng(numOrdenAnterior), claveRes)
    End If

    Dim pwdHojas As String
    pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()

    ' --- 1. GESTIÓN DE PROTECCIÓN EN HOJA DE RESIDENCIA (ws) ---
    Dim estabaProtegidaWs As Boolean
    estabaProtegidaWs = ws.ProtectContents
    If estabaProtegidaWs Then
        On Error Resume Next
        If Len(pwdHojas) > 0 Then ws.Unprotect Password:=pwdHojas
        If ws.ProtectContents Then ws.Unprotect Password:=""
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo ErrorBorrado
    End If

    ' --- 2. BORRADO DE VALORES B:ultimaColumna (respetando fórmulas) ---
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

    ' --- 3. REPROTEGER HOJA DE RESIDENCIA SI ESTABA PROTEGIDA ---
    If estabaProtegidaWs Then
        On Error Resume Next
        ws.Protect Password:=pwdHojas, UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
            AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
        On Error GoTo ErrorBorrado
    End If

    ' --- 4. BORRADO DE FILA EN HOJA RESUMEN (con desprotección segura) ---
    Call BorrarFilaEnResumen_Central(numOrdenAnterior, nombreResumen, pwdHojas)

    GoTo Restaurar

ErrorBorrado:
    ' Asegurar reprotección de residencia ante cualquier error inesperado
    If estabaProtegidaWs And Not ws Is Nothing Then
        On Error Resume Next
        ws.Protect Password:=pwdHojas, UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
            AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
        On Error GoTo 0
    End If
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "BorradoNumOrden"

Restaurar:
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic

End Sub


' =============================================================================
' PRIVADA: BorrarFilaEnResumen_Central
' Elimina la fila de la orden en la hoja RESUMEN desprotegiendo y reprotegiendo
' de forma transparente.
' =============================================================================
Private Sub BorrarFilaEnResumen_Central( _
    ByVal numOrden As Variant, _
    ByVal nombreHojaResumen As String, _
    ByVal pwdHojas As String)

    Dim wsResumen As Worksheet
    Dim celdaEncontrada As Range
    Dim ultimaFila As Long
    Dim numOrdenStr As String

    numOrdenStr = Trim(CStr(numOrden))
    If numOrdenStr = "" Then Exit Sub

    ' Obtención segura de la hoja (inmune a tildes/codepages)
    On Error Resume Next
    Set wsResumen = modDatabase.ObtenerHojaSegura(nombreHojaResumen)
    If wsResumen Is Nothing Then
        Set wsResumen = ThisWorkbook.Worksheets(nombreHojaResumen)
    End If
    On Error GoTo 0

    If wsResumen Is Nothing Then Exit Sub

    ultimaFila = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    If ultimaFila < 2 Then Exit Sub

    ' Buscar el Nº de Orden en columna A del resumen
    On Error Resume Next
    Set celdaEncontrada = wsResumen.Range("A2:A" & ultimaFila).Find( _
        What:=numOrdenStr, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    On Error GoTo 0

    If Not celdaEncontrada Is Nothing Then
        ' Desproteger RESUMEN temporalmente para poder eliminar la fila entera
        Dim estabaProtegidaResumen As Boolean
        estabaProtegidaResumen = wsResumen.ProtectContents
        
        If estabaProtegidaResumen Then
            On Error Resume Next
            If Len(pwdHojas) > 0 Then wsResumen.Unprotect Password:=pwdHojas
            If wsResumen.ProtectContents Then wsResumen.Unprotect Password:=""
            If wsResumen.ProtectContents Then wsResumen.Unprotect
            On Error GoTo 0
        End If

        ' Eliminar la fila
        celdaEncontrada.EntireRow.Delete

        ' Reproteger RESUMEN restaurando sus permisos
        If estabaProtegidaResumen Then
            On Error Resume Next
            wsResumen.Protect Password:=pwdHojas, UserInterfaceOnly:=True, _
                AllowFormattingCells:=True, AllowFiltering:=True, AllowSorting:=True
            On Error GoTo 0
        End If
    End If

End Sub
