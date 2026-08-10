Attribute VB_Name = "SolapesAceptadosOviedo"
'Attribute VB_Name = "SolapesAceptadosOviedo"
Option Explicit

'==========================================================
' SolapesAceptadosOviedo
' - Lista blanca por GRUPO: (OVIEDO + Habitaci�n) y rango fechas
' - Persistencia en hoja VeryHidden: SOLAPES_ACEPTADOS_OVIEDO
'==========================================================

Private Const SH_SOLAPES_OK As String = "SOLAPES_ACEPTADOS_OVIEDO"

' Columnas (1-based)
Private Const COL_CLAVE  As Long = 1  ' A: Clave
Private Const COL_DESDE  As Long = 2  ' B: FechaDesde
Private Const COL_HASTA  As Long = 3  ' C: FechaHasta
Private Const COL_MOTIVO As Long = 4  ' D: Motivo
Private Const COL_TS     As Long = 5  ' E: Timestamp
Private Const COL_USER   As Long = 6  ' F: Usuario

'========================
' API PUBLICA
'========================

Public Function ClaveSolapeGrupoOviedo(ByVal habStr As String) As String
    ClaveSolapeGrupoOviedo = "OVIEDO|HAB_" & Trim$(habStr) & "|GRUPO"
End Function

Public Function CargarSolapesAceptadosOviedo() As Object
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHojaSolapesOK()

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_CLAVE).End(xlUp).Row
    If lastRow < 2 Then
        Set CargarSolapesAceptadosOviedo = dict
        Exit Function
    End If

    Dim r As Long
    For r = 2 To lastRow
        Dim clave As String
        clave = Trim(CStr(ws.Cells(r, COL_CLAVE).Value))

        If clave <> "" And IsDate(ws.Cells(r, COL_DESDE).Value) And IsDate(ws.Cells(r, COL_HASTA).Value) Then
            Dim d1 As Date, d2 As Date
            d1 = CDate(ws.Cells(r, COL_DESDE).Value)
            d2 = CDate(ws.Cells(r, COL_HASTA).Value)

            If d2 < d1 Then
                Dim tmp As Date: tmp = d1: d1 = d2: d2 = tmp
            End If

            If Not dict.Exists(clave) Then
                Dim col As Collection
                Set col = New Collection
                dict.Add clave, col
            End If

            Dim tramo(1 To 2) As Date
            tramo(1) = d1
            tramo(2) = d2
            dict(clave).Add tramo
        End If
    Next r

    Set CargarSolapesAceptadosOviedo = dict
End Function

Public Function SolapeAceptadoOviedo(ByVal dict As Object, ByVal claveSolape As String, ByVal fechaSolape As Date) As Boolean
    SolapeAceptadoOviedo = False
    If dict Is Nothing Then Exit Function
    If Trim$(claveSolape) = "" Then Exit Function
    If Not dict.Exists(claveSolape) Then Exit Function

    Dim col As Collection
    Set col = dict(claveSolape)

    Dim i As Long
    For i = 1 To col.Count
        Dim tramo As Variant
        tramo = col(i)
        If fechaSolape >= CDate(tramo(1)) And fechaSolape <= CDate(tramo(2)) Then
            SolapeAceptadoOviedo = True
            Exit Function
        End If
    Next i
End Function

Public Sub GuardarSolapeAceptadoOviedo(ByVal claveSolape As String, ByVal fechaDesde As Date, ByVal fechaHasta As Date, Optional ByVal motivo As String = "")
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHojaSolapesOK()

    If fechaHasta < fechaDesde Then
        Dim tmp As Date: tmp = fechaDesde: fechaDesde = fechaHasta: fechaHasta = tmp
    End If

    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.Count, COL_CLAVE).End(xlUp).Row + 1
    If nextRow < 2 Then nextRow = 2

    ws.Cells(nextRow, COL_CLAVE).Value = claveSolape
    ws.Cells(nextRow, COL_DESDE).Value = fechaDesde
    ws.Cells(nextRow, COL_HASTA).Value = fechaHasta
    ws.Cells(nextRow, COL_MOTIVO).Value = motivo
    ws.Cells(nextRow, COL_TS).Value = Now

    Dim u As String
    On Error Resume Next
    u = Trim$(CStr(ThisWorkbook.usuarioActual))
    On Error GoTo 0
    If u = "" Then u = "(sin login)"
    ws.Cells(nextRow, COL_USER).Value = u
End Sub

'========================
' INTERNA
'========================

Private Function ObtenerOCrearHojaSolapesOK() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SH_SOLAPES_OK)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SH_SOLAPES_OK

        ws.Cells(1, COL_CLAVE).Value = "Clave"
        ws.Cells(1, COL_DESDE).Value = "FechaDesde"
        ws.Cells(1, COL_HASTA).Value = "FechaHasta"
        ws.Cells(1, COL_MOTIVO).Value = "Motivo"
        ws.Cells(1, COL_TS).Value = "Timestamp"
        ws.Cells(1, COL_USER).Value = "Usuario"

        ws.Rows(1).Font.Bold = True
        ws.Columns(COL_DESDE).NumberFormat = "dd/mm/yyyy"
        ws.Columns(COL_HASTA).NumberFormat = "dd/mm/yyyy"

        ws.visible = xlSheetVeryHidden
    End If

    Set ObtenerOCrearHojaSolapesOK = ws
End Function
