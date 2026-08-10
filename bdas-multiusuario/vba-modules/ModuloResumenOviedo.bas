Attribute VB_Name = "ModuloResumenOviedo"
'=================================================================================
' M�dulo: ModuloResumenOviedo v14.11.3
' Versi�n: 3.0
' Fecha: 2025-03-17
' Prop�sito: Actualizar autom�ticamente RESUMEN OVIEDO eliminando filas vac�as
'           (N� ORDEN que ya no existen en RESIDENCIA OVIEDO)
'=================================================================================
Option Explicit

'=================================================================================
' SUB: ActualizarResumenOviedo (OPTIMIZADA v3.0 - Arrays en memoria)
'=================================================================================
Public Sub ActualizarResumenOviedo(Optional esArranque As Boolean = False)
    On Error GoTo ErrorHandler
    
        Dim wsResidencia As Worksheet
    Dim wsResumen As Worksheet
    Dim estabaProtegida As Boolean
    Dim ultimaFilaResidencia As Long
    Dim ultimaFilaResumen As Long
    
    On Error Resume Next
    Set wsResidencia = ThisWorkbook.Worksheets("RESIDENCIA OVIEDO")
    Set wsResumen = ThisWorkbook.Worksheets("RESUMEN OVIEDO")
    On Error GoTo ErrorHandler
    
    If wsResidencia Is Nothing Or wsResumen Is Nothing Then Exit Sub
    
    ' === LIMPIAR FILTROS ACTIVA PARA EVITAR C�LCULOS ERR�NEOS Y ERRORES DE ESCRITURA ===
    If wsResidencia.FilterMode Then wsResidencia.ShowAllData
    If wsResumen.FilterMode Then wsResumen.ShowAllData
    
    ' === DESPROTEGER HOJA TEMPORALMENTE ===
    estabaProtegida = wsResumen.ProtectContents
    If estabaProtegida Then wsResumen.Unprotect password:=""
    
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    ultimaFilaResidencia = wsResidencia.Cells(wsResidencia.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFilaResidencia < 2 Then
        ultimaFilaResumen = wsResumen.Cells(wsResumen.Rows.Count, "A").End(xlUp).Row
        If ultimaFilaResumen >= 2 Then
            wsResumen.Range("A2:A" & ultimaFilaResumen).ClearContents
            wsResumen.Range("L2:O" & ultimaFilaResumen).ClearContents
        End If
        GoTo CleanExit
    End If
    
    ' === PASO 1: Leer RESIDENCIA en arrays ===
    Dim arrOrdenesRes As Variant
    Dim arrPagadoRes As Variant
    Dim arrFacturaRes As Variant
    
    arrOrdenesRes = wsResidencia.Range("A2:A" & ultimaFilaResidencia).Value
    arrPagadoRes = wsResidencia.Range("AB2:AB" & ultimaFilaResidencia).Value
    arrFacturaRes = wsResidencia.Range("C2:C" & ultimaFilaResidencia).Value
    
    ' Protecci�n: rango de una sola fila devuelve escalar
    If Not IsArray(arrOrdenesRes) Then
        Dim arrTempO3(1 To 1, 1 To 1) As Variant
        Dim arrTempP3(1 To 1, 1 To 1) As Variant
        Dim arrTempF3(1 To 1, 1 To 1) As Variant
        arrTempO3(1, 1) = arrOrdenesRes
        arrTempP3(1, 1) = arrPagadoRes
        arrTempF3(1, 1) = arrFacturaRes
        arrOrdenesRes = arrTempO3
        arrPagadoRes = arrTempP3
        arrFacturaRes = arrTempF3
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
            wsResumen.Range("L2:O" & ultimaFilaResumen).ClearContents
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
    
    ' === PASO 3: Limpiar hu�rfanos L-O en memoria ===
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
            arrResumenActual = Array(arrResumenActual)
            Dim arrTemp(1 To 1, 1 To 1) As Variant
            arrTemp(1, 1) = arrResumenActual(0)
            arrResumenActual = arrTemp
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
    
        If ultimaFilaResumen > contadorValidas + 1 Then
        wsResumen.Range("M" & contadorValidas + 2 & ":N" & ultimaFilaResumen).ClearContents
    End If
    
    ' === PROPAGAR F�RMULAS EN LAS COLUMNAS CALCULADAS (NUEVO) ===
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
    ' Restaurar protecci�n si estaba protegida
    If estabaProtegida And Not wsResumen Is Nothing Then
        wsResumen.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    Exit Sub

ErrorHandler:
    ' Restaurar protecci�n en caso de error
    If estabaProtegida And Not wsResumen Is Nothing Then
        wsResumen.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True, AllowFiltering:=True, AllowSorting:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "Error al actualizar RESUMEN OVIEDO: " & Err.Description, vbCritical, "Error"
End Sub
