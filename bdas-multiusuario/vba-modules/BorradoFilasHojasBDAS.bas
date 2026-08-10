Attribute VB_Name = "BorradoFilasHojasBDAS"
Option Explicit

' ==============================================================================
' M�DULO: BorradoFilasHojasBDAS
' ------------------------------------------------------------------------------
' Elimina de forma definitiva las filas de las hojas BDAS cuyo campo
' "FECHA ARCHIVADO" (columna M) tenga una antig�edad igual o superior
' al n�mero de d�as configurado( 3 A�OS ).
'
' - Hojas afectadas:
'       � BDAS GIJ�N
'       � BDAS SOTO
'       � BDAS OVIEDO
'
' - Columna de referencia:
'       � Columna M (FECHA ARCHIVADO)
'
' - El borrado es f�sico (Delete Row):
'       � No quedan filas en blanco
'       � El resto de registros se desplazan hacia arriba
'
' - Pensado para ejecutarse autom�ticamente desde Workbook_Open
' - Rendimiento optimizado (sin impacto perceptible)
' ==============================================================================

Public Sub BorrarFilasBDAS()
    Const DIAS_MAX As Long = 1095
    Dim hojasBDAS As Variant
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim fechaArchivado As Variant
    Dim filasEliminadas As Long
    Dim mensaje As String
    Dim nombreHoja As Variant
    Dim tempWs As Worksheet
    Dim rngEliminar As Range
    hojasBDAS = Array("BDAS GIJ�N", "BDAS SOTO", "BDAS OVIEDO")
    On Error GoTo SalidaSegura
    For Each nombreHoja In hojasBDAS
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(nombreHoja)
        On Error GoTo 0
        
        If Not ws Is Nothing Then
            ' OPTIMIZACI�N: Solo depurar si la residencia correspondiente est� visible
            Dim nombreResidencia As String
            nombreResidencia = Replace(ws.Name, "BDAS", "RESIDENCIA")
            
            Dim resVisible As Boolean
            resVisible = False
            
            Set tempWs = Nothing
            On Error Resume Next
            Set tempWs = ThisWorkbook.Worksheets(nombreResidencia)
            On Error GoTo 0
            
            If Not tempWs Is Nothing Then
                If tempWs.visible = xlSheetVisible Then
                    resVisible = True
                End If
            End If
            
            If resVisible Then
                filasEliminadas = 0
                ultimaFila = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row
                If ultimaFila <= 1 Then GoTo SiguienteHoja
                Set rngEliminar = Nothing
                
                For i = ultimaFila To 2 Step -1
                    fechaArchivado = ws.Cells(i, "M").Value
                    If IsDate(fechaArchivado) Then
                        If Date - CDate(fechaArchivado) >= DIAS_MAX Then
                            If rngEliminar Is Nothing Then
                                Set rngEliminar = ws.Rows(i)
                            Else
                                Set rngEliminar = Union(rngEliminar, ws.Rows(i))
                            End If
                            filasEliminadas = filasEliminadas + 1
                        End If
                    End If
                Next i
                
                If Not rngEliminar Is Nothing Then
                    rngEliminar.Delete
                End If
            End If
        End If
SiguienteHoja:
    Next nombreHoja
    If mensaje <> "" Then
        MsgBox "Se han eliminado definitivamente registros archivados:" & _
               vbCrLf & vbCrLf & mensaje, _
               vbInformation, "Depuraci�n autom�tica BDAS"
    End If
SalidaSegura:
End Sub

