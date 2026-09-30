Attribute VB_Name = "ModuloResumenSoto"
'=================================================================================
' Módulo: ModuloResumenSoto
' Versión: 3.0
' Fecha: 2025-03-17
' Propósito: Actualizar automáticamente RESUMEN SOTO eliminando filas vacías
'           (Nº ORDEN que ya no existen en RESIDENCIA SOTO)
'=================================================================================
Option Explicit

'=================================================================================
' SUB: ActualizarResumenSoto (OPTIMIZADA v3.0 - Arrays en memoria)
'=================================================================================
Public Sub ActualizarResumenSoto(Optional esArranque As Boolean = False)
    On Error GoTo ErrorHandler
    
        Dim wsResidencia As Worksheet
    Dim wsResumen As Worksheet
    Dim estabaProtegida As Boolean
    Dim ultimaFilaResidencia As Long
    Dim ultimaFilaResumen As Long
    
    On Error Resume Next
    Set wsResidencia = ThisWorkbook.Worksheets("RESIDENCIA SOTO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN SOTO")
    On Error GoTo ErrorHandler
    
    If wsResidencia Is Nothing Or wsResumen Is Nothing Then Exit Sub
    
    ' === DESPROTEGER HOJA TEMPORALMENTE (ANTES DE LIMPIAR FILTROS) ===
    Dim pwdHojas As String
    pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
    
    estabaProtegida = wsResumen.ProtectContents
    If estabaProtegida Then
        On Error Resume Next
        If Len(pwdHojas) > 0 Then wsResumen.Unprotect password:=pwdHojas
        If wsResumen.ProtectContents Then wsResumen.Unprotect password:=""
        If wsResumen.ProtectContents Then wsResumen.Unprotect
        On Error GoTo ErrorHandler
    End If
    
    ' === LIMPIAR FILTROS DE FORMA SEGURA CON LA HOJA YA DESPROTEGIDA ===
    On Error Resume Next
    If wsResidencia.FilterMode Then wsResidencia.ShowAllData
    If wsResumen.FilterMode Then wsResumen.ShowAllData
    On Error GoTo ErrorHandler
    
    Application.screenUpdating = False

    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    ultimaFilaResidencia = wsResidencia.Cells(wsResidencia.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaResidencia < 2 Then
        ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
        If ultimaFilaResumen >= 2 Then
            wsResumen.Range("A2:A" & ultimaFilaResumen).ClearContents
            wsResumen.Range("L2:N" & ultimaFilaResumen).ClearContents
        End If
        GoTo CleanExit
    End If
    
    ' === PASO 1: Leer RESIDENCIA en arrays ===
    Dim arrOrdenesRes As Variant
    Dim arrPagadoRes As Variant
    Dim arrFacturaRes As Variant
    
    arrOrdenesRes = wsResidencia.Range("A2:A" & ultimaFilaResidencia).Value
    arrPagadoRes = wsResidencia.Range("AA2:AA" & ultimaFilaResidencia).Value
    arrFacturaRes = wsResidencia.Range("C2:C" & ultimaFilaResidencia).Value
    
    ' Protección: rango de una sola fila devuelve escalar
    If Not IsArray(arrOrdenesRes) Then
        Dim arrTempO2(1 To 1, 1 To 1) As Variant
        Dim arrTempP2(1 To 1, 1 To 1) As Variant
        Dim arrTempF2(1 To 1, 1 To 1) As Variant
        arrTempO2(1, 1) = arrOrdenesRes
        arrTempP2(1, 1) = arrPagadoRes
        arrTempF2(1, 1) = arrFacturaRes
        arrOrdenesRes = arrTempO2
        arrPagadoRes = arrTempP2
        arrFacturaRes = arrTempF2
    End If
    
    ' === PASO 2: Construir arrays de salida ===
    Dim totalFilas As Long
    totalFilas = UBound(arrOrdenesRes, 1)
    
    Dim contadorValidas As Long
    Dim i As Long
    contadorValidas = 0
    
    For i = 1 To totalFilas
        If Trim(CStr(arrOrdenesRes(i, 1))) <> "" Then
            contadorValidas = contadorValidas + 1
        End If
    Next i
    
    If contadorValidas = 0 Then
        ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
        If ultimaFilaResumen >= 2 Then
            wsResumen.Range("A2:A" & ultimaFilaResumen).ClearContents
            wsResumen.Range("L2:N" & ultimaFilaResumen).ClearContents
        End If
        GoTo CleanExit
    End If
    
    Dim arrOrdenesSalida() As Variant
    Dim arrPagadoSalida() As Variant
    Dim arrFacturaSalida() As Variant
    
    ReDim arrOrdenesSalida(1 To contadorValidas, 1 To 1)
    ReDim arrPagadoSalida(1 To contadorValidas, 1 To 1)
    ReDim arrFacturaSalida(1 To contadorValidas, 1 To 1)
    
    Dim idx As Long
    idx = 0
    
    For i = 1 To totalFilas
        If Trim(CStr(arrOrdenesRes(i, 1))) <> "" Then
            idx = idx + 1
            arrOrdenesSalida(idx, 1) = arrOrdenesRes(i, 1)
            arrPagadoSalida(idx, 1) = arrPagadoRes(i, 1)
            arrFacturaSalida(idx, 1) = arrFacturaRes(i, 1)
        End If
    Next i
    
    ' === PASO 3: Limpiar huérfanos L-N en memoria ===
    ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaResumen >= 2 Then
        Dim dictOrdenesActuales As Object
        Set dictOrdenesActuales = CreateObject("Scripting.Dictionary")
        
        For i = 1 To contadorValidas
            dictOrdenesActuales(CStr(arrOrdenesSalida(i, 1))) = True
        Next i
        
        Dim arrResumenActual As Variant
        arrResumenActual = wsResumen.Range("A2:A" & ultimaFilaResumen).Value
        
        If Not IsArray(arrResumenActual) Then
            Dim arrTempS(1 To 1, 1 To 1) As Variant
            arrTempS(1, 1) = arrResumenActual
            arrResumenActual = arrTempS
        End If
        
        ' Lectura y limpieza optimizada en memoria
        Dim arrLN As Variant
        arrLN = wsResumen.Range("L2:N" & ultimaFilaResumen).Value
        Dim modificado As Boolean
        modificado = False
        
        If IsArray(arrLN) Then
            For i = 1 To UBound(arrResumenActual, 1)
                If Trim(CStr(arrResumenActual(i, 1))) <> "" Then
                    If Not dictOrdenesActuales.Exists(CStr(arrResumenActual(i, 1))) Then
                        arrLN(i, 1) = Empty
                        arrLN(i, 2) = Empty
                        arrLN(i, 3) = Empty
                        modificado = True
                    End If
                End If
            Next i
            
            If modificado Then
                wsResumen.Range("L2:N" & ultimaFilaResumen).Value = arrLN
            End If
        End If
        
        Set dictOrdenesActuales = Nothing
    End If
    
    ' === PASO 4: Escribir todo de una sola vez ===
    Dim maxFilaLimpiar As Long
    maxFilaLimpiar = Application.WorksheetFunction.Max(ultimaFilaResumen, contadorValidas + 1)
    
    If maxFilaLimpiar >= 2 Then
        wsResumen.Range("A2:A" & maxFilaLimpiar).ClearContents
    End If
    
    ' SOTO: A=Órdenes, L=Pagado, M=Factura
    wsResumen.Range("A2").Resize(contadorValidas, 1).Value = arrOrdenesSalida
    wsResumen.Range("L2").Resize(contadorValidas, 1).Value = arrPagadoSalida
    wsResumen.Range("M2").Resize(contadorValidas, 1).Value = arrFacturaSalida
    
        If ultimaFilaResumen > contadorValidas + 1 Then
        wsResumen.Range("L" & contadorValidas + 2 & ":M" & ultimaFilaResumen).ClearContents
    End If
    
    ' === PROPAGAR FÓRMULAS EN LAS COLUMNAS CALCULADAS (NUEVO) ===
    Dim col As Long
    Dim uCol As Long
    If contadorValidas > 1 Then
        uCol = wsResumen.Cells(2, wsResumen.Columns.Count).End(xlToLeft).Column
        For col = 1 To uCol
            If wsResumen.Cells(2, col).HasFormula Then
                wsResumen.Cells(2, col).AutoFill _
                    Destination:=wsResumen.Range(wsResumen.Cells(2, col), wsResumen.Cells(contadorValidas + 1, col)), _
                    Type:=xlFillDefault
            End If
        Next col
    End If
    
CleanExit:
    ' Restaurar protección si estaba protegida
    If estabaProtegida And Not wsResumen Is Nothing Then
        On Error Resume Next
        wsResumen.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), _
            UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True, AllowSorting:=True
        On Error GoTo 0
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    Exit Sub

ErrorHandler:
    ' Restaurar protección en caso de error
    If estabaProtegida And Not wsResumen Is Nothing Then
        On Error Resume Next
        wsResumen.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), _
            UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True, AllowSorting:=True
        On Error GoTo 0
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "Error al actualizar RESUMEN SOTO: " & Err.Description, vbCritical, "Error"
End Sub

