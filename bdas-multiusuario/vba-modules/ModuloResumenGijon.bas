Attribute VB_Name = "ModuloResumenGijon"
'=================================================================================
' Módulo: ModuloResumenGijon
' Versión: 3.0
' Fecha: 2026-03-17
' Propósito: Actualizar automáticamente RESUMEN GIJÓN eliminando filas vacías
'           (Nº ORDEN que ya no existen en RESIDENCIA GIJÓN)
'=================================================================================
Option Explicit

'=================================================================================
' SUB: ActualizarResumenGijon
' Actualiza la columna A de RESUMEN GIJÓN copiando solo los Nº ORDEN que
' existen actualmente en RESIDENCIA GIJÓN (sin huecos)
'=================================================================================
'=================================================================================
' SUB: ActualizarResumenGijon (OPTIMIZADA v3.0 - Arrays en memoria)
'=================================================================================
Public Sub ActualizarResumenGijon(Optional esArranque As Boolean = False)
    On Error GoTo ErrorHandler
    
        Dim wsResidencia As Worksheet
    Dim wsResumen As Worksheet
    Dim estabaProtegida As Boolean
    Dim ultimaFilaResidencia As Long
    Dim ultimaFilaResumen As Long
    
    On Error Resume Next
    Set wsResidencia = ThisWorkbook.Worksheets("RESIDENCIA GIJÓN")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN GIJÓN")
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
    
    ' --- Si no hay datos en RESIDENCIA, limpiar RESUMEN y salir ---
    If ultimaFilaResidencia < 2 Then
        ultimaFilaResumen = Application.WorksheetFunction.Max( _
            wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row, _
            wsResumen.Cells(wsResumen.Rows.Count, "B").End(xlUp).Row)
        If ultimaFilaResumen >= 2 Then
            wsResumen.Range("A2:A" & ultimaFilaResumen).ClearContents
            wsResumen.Range("L2:O" & ultimaFilaResumen).ClearContents
        End If
        GoTo CleanExit
    End If
    
    ' === PASO 1: Leer RESIDENCIA completa en arrays ===
    Dim arrOrdenesRes As Variant    ' Columna A (Nº ORDEN)
    Dim arrPagadoRes As Variant     ' Columna AA (PAGADO)
    Dim arrFacturaRes As Variant    ' Columna C (Nº FACTURA)
    
    arrOrdenesRes = wsResidencia.Range("A2:A" & ultimaFilaResidencia).Value
    arrPagadoRes = wsResidencia.Range("AA2:AA" & ultimaFilaResidencia).Value
    arrFacturaRes = wsResidencia.Range("C2:C" & ultimaFilaResidencia).Value
    
    ' Protección: rango de una sola fila devuelve escalar
    If Not IsArray(arrOrdenesRes) Then
        Dim arrTempO1(1 To 1, 1 To 1) As Variant
        Dim arrTempP1(1 To 1, 1 To 1) As Variant
        Dim arrTempF1(1 To 1, 1 To 1) As Variant
        arrTempO1(1, 1) = arrOrdenesRes
        arrTempP1(1, 1) = arrPagadoRes
        arrTempF1(1, 1) = arrFacturaRes
        arrOrdenesRes = arrTempO1
        arrPagadoRes = arrTempP1
        arrFacturaRes = arrTempF1
    End If
    
    ' === PASO 2: Construir arrays de salida en memoria ===
    Dim totalFilas As Long
    totalFilas = UBound(arrOrdenesRes, 1)
    
    ' Contar órdenes válidas
    Dim contadorValidas As Long
    Dim i As Long
    contadorValidas = 0
    
    For i = 1 To totalFilas
        If Trim(CStr(arrOrdenesRes(i, 1))) <> "" Then
            contadorValidas = contadorValidas + 1
        End If
    Next i
    
    ' Si no hay órdenes válidas, limpiar y salir
    If contadorValidas = 0 Then
        ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
        If ultimaFilaResumen >= 2 Then
            wsResumen.Range("A2:A" & ultimaFilaResumen).ClearContents
            wsResumen.Range("L2:O" & ultimaFilaResumen).ClearContents
        End If
        GoTo CleanExit
    End If
    
    ' Crear arrays de salida
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
    
    ' === PASO 3: Limpiar huérfanos en columnas L-O en memoria ===
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
            Dim arrTempG(1 To 1, 1 To 1) As Variant
            arrTempG(1, 1) = arrResumenActual
            arrResumenActual = arrTempG
        End If
        
        ' Lectura y limpieza optimizada en memoria
        Dim arrLO As Variant
        arrLO = wsResumen.Range("L2:O" & ultimaFilaResumen).Value
        Dim modificado As Boolean
        modificado = False
        
        If IsArray(arrLO) Then
            For i = 1 To UBound(arrResumenActual, 1)
                If Trim(CStr(arrResumenActual(i, 1))) <> "" Then
                    If Not dictOrdenesActuales.Exists(CStr(arrResumenActual(i, 1))) Then
                        arrLO(i, 1) = Empty
                        arrLO(i, 2) = Empty
                        arrLO(i, 3) = Empty
                        arrLO(i, 4) = Empty
                        modificado = True
                    End If
                End If
            Next i
            
            If modificado Then
                wsResumen.Range("L2:O" & ultimaFilaResumen).Value = arrLO
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
    
    wsResumen.Range("A2").Resize(contadorValidas, 1).Value = arrOrdenesSalida
    wsResumen.Range("M2").Resize(contadorValidas, 1).Value = arrPagadoSalida
    wsResumen.Range("N2").Resize(contadorValidas, 1).Value = arrFacturaSalida
    
    ' === PROPAGAR FÓRMULAS EN LAS COLUMNAS CALCULADAS ===
    Dim col As Long
    Dim uCol As Long
    uCol = wsResumen.Cells(2, wsResumen.Columns.Count).End(xlToLeft).Column
    
    If contadorValidas > 1 Then
        For col = 1 To uCol
            If wsResumen.Cells(2, col).HasFormula Then
                wsResumen.Cells(2, col).AutoFill _
                    Destination:=wsResumen.Range(wsResumen.Cells(2, col), wsResumen.Cells(contadorValidas + 1, col)), _
                    Type:=xlFillDefault
            End If
        Next col
    End If
    
    ' Limpiar fórmulas y valores huérfanos más allá del último registro válido
    If ultimaFilaResumen > contadorValidas + 1 Then
        wsResumen.Range(wsResumen.Cells(contadorValidas + 2, 2), wsResumen.Cells(ultimaFilaResumen, uCol)).ClearContents
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
    MsgBox "Error al actualizar RESUMEN GIJÓN: " & Err.Description, vbCritical, "Error"
End Sub


