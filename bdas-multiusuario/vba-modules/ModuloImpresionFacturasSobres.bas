Attribute VB_Name = "ModuloImpresionFacturasSobres"
Option Explicit

'==============================================================================
' M�DULO: ModuloImpresionFacturasSobres v14.11.3
'
' DESCRIPCI�N:
'   M�dulo centralizado para la impresi�n de SOBRES y FACTURAS desde la hoja
'   "RESUMEN", con validaciones completas y control de errores.
'
' MAPA DE COLUMNAS (HOJA RESUMEN):
'   A  -> N� ORDEN
'   B  -> ESTADO (SI / CONCEDIDA / REEVALUADA / etc.)
'   L  -> SOBRE   (marca "IMPRESO")
'   N  -> N� FACTURA
'
' NOTA:
'   Las columnas L y N han sido invertidas respecto a versiones anteriores.
'==============================================================================

'---------------------------
' CONSTANTES DE COLUMNAS
'---------------------------
Private Const COL_ORDEN    As String = "A"
Private Const COL_ESTADO   As String = "B"
Private Const COL_SOBRE    As String = "L"   ' << AHORA SOBRE
Private Const COL_FACTURA As String = "N"   ' << AHORA FACTURA

'---------------------------
' CONSTANTES GENERALES
'---------------------------
Private Const MARCA_IMPRESO As String = "IMPRESO"

'==============================================================================
' SeleccionImprimirSobre
'
' - Se invoca al seleccionar la columna SOBRE (L).
' - Valida el estado de la reserva.
' - Pregunta al usuario si desea imprimir el sobre.
' - Marca "IMPRESO" y ejecuta la impresi�n.
'==============================================================================
Public Sub SeleccionImprimirSobre( _
    hojaResumen As Worksheet, _
    fila As Long, _
    hojaDestinoNombre As String, _
    celdaDestino As String, _
    Optional estadoValido1 As String = "SI", _
    Optional estadoValido2 As String = "CONCEDIDA", _
    Optional estadoValido3 As String = "REEVALUADA" _
)
    On Error GoTo CleanExit
    Application.enableEvents = False

    Dim estado As String
    estado = UCase$(Trim$(CStr(hojaResumen.Cells(fila, COL_ESTADO).Value)))

    Dim nroOrden As Variant
    nroOrden = hojaResumen.Cells(fila, COL_ORDEN).Value

    ' Validaci�n de estado
    If estado = estadoValido1 Or estado = estadoValido2 Or estado = estadoValido3 Then
        If MsgBox("�Desea imprimir el sobre del N� ORDEN " & nroOrden & "?", _
                  vbYesNo + vbQuestion, "Imprimir Sobre") = vbYes Then

            hojaResumen.Cells(fila, COL_SOBRE).Value = MARCA_IMPRESO

            Call ImprimirSobre( _
                hojaResumen, fila, hojaDestinoNombre, celdaDestino, _
                estadoValido1, estadoValido2, estadoValido3 _
            )
        End If
    End If

CleanExit:
    Application.enableEvents = True
End Sub

'==============================================================================
' ImprimirSobre
'
' - Verifica que la columna SOBRE (L) tenga la marca "IMPRESO".
' - Revalida el estado.
' - Copia el N� ORDEN a la hoja destino y la activa.
'==============================================================================
Public Sub ImprimirSobre( _
    hojaResumen As Worksheet, _
    fila As Long, _
    hojaDestinoNombre As String, _
    celdaDestino As String, _
    Optional estadoValido1 As String = "SI", _
    Optional estadoValido2 As String = "CONCEDIDA", _
    Optional estadoValido3 As String = "REEVALUADA" _
)
    On Error GoTo CleanExit
    Application.enableEvents = False

    Dim marcaSobre As String
    marcaSobre = UCase$(Trim$(CStr(hojaResumen.Cells(fila, COL_SOBRE).Value)))
    If marcaSobre <> MARCA_IMPRESO Then GoTo CleanExit

    Dim estado As String
    estado = UCase$(Trim$(CStr(hojaResumen.Cells(fila, COL_ESTADO).Value)))

    If Not (estado = estadoValido1 Or estado = estadoValido2 Or estado = estadoValido3) Then
        MsgBox "No procede la impresi�n del sobre para esta reserva.", vbInformation
        hojaResumen.Cells(fila, COL_SOBRE).ClearContents
        GoTo CleanExit
    End If

    Dim wsDestino As Worksheet
    Set wsDestino = ObtenerHoja(hojaResumen.Parent, hojaDestinoNombre)
    If wsDestino Is Nothing Then GoTo CleanExit

        ' === ESCRIBIR EL VALOR Y FORZAR C�LCULOS (OPTIMIZADO) ===
    ' Desactivamos eventos solo durante la escritura del valor para evitar loops
    Application.enableEvents = False
    wsDestino.Range(celdaDestino).Value = hojaResumen.Cells(fila, COL_ORDEN).Value
    Application.enableEvents = True
    
    ' Activamos la hoja y seleccionamos el destino
    wsDestino.Activate
    wsDestino.Range(celdaDestino).Select
    
    ' Recalculamos la hoja de destino de forma s�ncrona.
    ' Al estar enableEvents = True, esto dispara inmediatamente el Worksheet_Calculate
    ' de la hoja destino en memoria, sin retardos artificiales ni simulaci�n de teclado.
    wsDestino.Calculate

CleanExit:
    Application.enableEvents = True
End Sub

'==============================================================================
' ImprimirFactura
'
' - Se invoca al seleccionar la columna FACTURA (N).
' - Valida que el N� FACTURA sea num�rico e �ntegro.
' - Solicita confirmaci�n al usuario.
' - Copia el N� ORDEN a la hoja de impresi�n de factura.
'==============================================================================
Public Sub ImprimirFactura( _
    hojaResumen As Worksheet, _
    fila As Long, _
    hojaDestinoNombre As String, _
    celdaDestino As String, _
    Optional rangoMinimo As Double = 1, _
    Optional rangoMaximo As Double = 2000 _
)
    On Error GoTo CleanExit
    Application.enableEvents = False

    Dim nroFactura As Variant
    nroFactura = hojaResumen.Cells(fila, COL_FACTURA).Value

    If IsNumeric(nroFactura) _
       And nroFactura = Int(nroFactura) _
       And nroFactura >= rangoMinimo _
       And nroFactura <= rangoMaximo Then

        If MsgBox("�Desea imprimir la factura N� " & nroFactura & "?", _
                  vbYesNo + vbQuestion, "Imprimir Factura") = vbYes Then

            Dim wsDestino As Worksheet
            Set wsDestino = ObtenerHoja(hojaResumen.Parent, hojaDestinoNombre)
            If wsDestino Is Nothing Then GoTo CleanExit

            wsDestino.Range(celdaDestino).Value = hojaResumen.Cells(fila, COL_ORDEN).Value
            wsDestino.Activate
            wsDestino.Range(celdaDestino).Select
        End If
    End If

CleanExit:
    Application.enableEvents = True
End Sub

'==============================================================================
' ImprimirFacturaDesdeColumna
'
' Permite imprimir facturas desde una columna distinta a la est�ndar (N)
' Ej: SOTO usa la columna M
'==============================================================================
Public Sub ImprimirFacturaDesdeColumna( _
    hojaResumen As Worksheet, _
    fila As Long, _
    columnaFactura As String, _
    hojaDestinoNombre As String, _
    celdaDestino As String, _
    Optional rangoMinimo As Double = 1, _
    Optional rangoMaximo As Double = 2000 _
)
    On Error GoTo CleanExit
    Application.enableEvents = False

    Dim nroFactura As Variant
    nroFactura = hojaResumen.Cells(fila, columnaFactura).Value

    If IsNumeric(nroFactura) _
       And nroFactura = Int(nroFactura) _
       And nroFactura >= rangoMinimo _
       And nroFactura <= rangoMaximo Then

        If MsgBox("�Desea imprimir la factura N� " & nroFactura & "?", _
                  vbYesNo + vbQuestion, "Imprimir Factura") = vbYes Then

            Dim wsDestino As Worksheet
            Set wsDestino = ObtenerHoja(hojaResumen.Parent, hojaDestinoNombre)
            If wsDestino Is Nothing Then GoTo CleanExit

            wsDestino.Range(celdaDestino).Value = hojaResumen.Cells(fila, "A").Value
            wsDestino.Activate
            wsDestino.Range(celdaDestino).Select
        End If
    End If

CleanExit:
    Application.enableEvents = True
End Sub


'==============================================================================
' ObtenerHoja
'
' - Devuelve una hoja de forma segura
' - La hace visible si estaba oculta
'==============================================================================
Private Function ObtenerHoja(libro As Workbook, nombreHoja As String) As Worksheet
    On Error Resume Next
    Set ObtenerHoja = libro.Worksheets(nombreHoja)
    On Error GoTo 0

    If ObtenerHoja Is Nothing Then
        MsgBox "No existe la hoja '" & nombreHoja & "'.", vbExclamation
    ElseIf ObtenerHoja.visible <> xlSheetVisible Then
        ObtenerHoja.visible = xlSheetVisible
    End If
End Function

'==============================================================================
' VerificarIncoherenciasResidencia
'
' - Escanea la hoja en busca de incoherencias entre Resoluci�n y Habitaci�n.
' - Carga todos los datos en RAM para que la ejecuci�n sea INSTANT�NEA (< 2 ms).
'==============================================================================
Public Sub VerificarIncoherenciasResidencia(ws As Worksheet, colHab As String)
    On Error Resume Next ' Evita bloqueos ante celdas con errores de f�rmula (#N/A, etc.)
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub
    
    ' Obtener columna de habitaci�n num�rica
    Dim colHabIndex As Long
    If UCase(colHab) = "S" Then
        colHabIndex = 19 ' Columna S
    ElseIf UCase(colHab) = "T" Then
        colHabIndex = 20 ' Columna T
    Else
        Exit Sub
    End If
    
    ' OPTIMIZACI�N CLAVE: Leer todo el rango en una sola operaci�n a RAM
    Dim arrData As Variant
    arrData = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, colHabIndex)).Value
    
    Dim i As Long
    Dim nroOrden As String
    Dim resolucion As String
    Dim numHab As String
    Dim msg As String
    Dim conteo As Long
    
    conteo = 0
    msg = ""
    
    For i = 2 To lastRow
        nroOrden = Trim(CStr(arrData(i, 1))) ' Columna A (N� Orden)
        If nroOrden <> "" Then
            resolucion = UCase(Trim(CStr(arrData(i, 16)))) ' Columna P (Resoluci�n)
            numHab = Trim(CStr(arrData(i, colHabIndex)))     ' Columna S o T (Habitaci�n)
            
            Dim tieneResolucionPositiva As Boolean
            tieneResolucionPositiva = (resolucion = "SI" Or resolucion = "CONCEDIDA" Or resolucion = "REEVALUADA")
            
            Dim tieneHabitacion As Boolean
            tieneHabitacion = (numHab <> "")
            
            ' Si hay discrepancia
            If (tieneResolucionPositiva And Not tieneHabitacion) Or (Not tieneResolucionPositiva And tieneHabitacion) Then
                conteo = conteo + 1
                If conteo <= 10 Then
                    msg = msg & "- N� ORDEN " & nroOrden & " presenta incoherencia entre la RESOLUCI�N (" & IIf(resolucion = "", "vac�a", resolucion) & ") y N�M. HAB. (" & IIf(numHab = "", "vac�a", numHab) & ")." & vbCrLf
                End If
            End If
        End If
    Next i
    
    ' Mostrar advertencia si se encontraron desajustes
    If conteo > 0 Then
        If conteo > 10 Then
            msg = msg & "... y otras " & (conteo - 10) & " reservas m�s." & vbCrLf
        End If
        MsgBox "Se han detectado incoherencias en los datos grabados:" & vbCrLf & vbCrLf & msg & vbCrLf & "Por favor, revisa las resoluciones y habitaciones/apartamentos adjudicados.", vbExclamation + vbOKOnly, "Incoherencia de Datos"
    End If
End Sub
