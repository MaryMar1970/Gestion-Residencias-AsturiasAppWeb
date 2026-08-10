Attribute VB_Name = "DuplicadosOrdenyFactura"
Option Explicit
Public Function VerificarSecuenciaOrden(ws As Worksheet, _
                                        ByVal columna As String, _
                                        ByVal filaInicio As Long) As String
    Dim ultimaFila As Long
    Dim f As Long
    Dim valor As Variant
    Dim prevVal As Long
    Dim actual As Long
    
    ultimaFila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
    If ultimaFila < filaInicio Then Exit Function
    
    ' Buscar los dos �ltimos valores num�ricos de la columna
    ' para comprobar solo si hay hueco entre el pen�ltimo y el �ltimo
    Dim ultimoVal As Long
    Dim penultimoVal As Long
    Dim contadorNums As Long
    
    ultimoVal = 0
    penultimoVal = 0
    contadorNums = 0
    
    ' Recorrer de abajo hacia arriba para encontrar los 2 �ltimos n�meros
    For f = ultimaFila To filaInicio Step -1
        valor = ws.Cells(f, columna).Value
        If IsNumeric(valor) And CLng(valor) > 0 Then
            contadorNums = contadorNums + 1
            If contadorNums = 1 Then
                ultimoVal = CLng(valor)
            ElseIf contadorNums = 2 Then
                penultimoVal = CLng(valor)
                Exit For
            End If
        End If
    Next f
    
    ' Si no hay al menos 2 n�meros, no hay nada que comparar
    If contadorNums < 2 Then Exit Function
    
    ' Si el �ltimo es menor que el pen�ltimo ? reinicio de numeraci�n, no es error
    If ultimoVal < penultimoVal Then Exit Function
    
    ' Si hay hueco entre el pen�ltimo y el �ltimo ? avisar
    If ultimoVal > penultimoVal + 1 Then
        VerificarSecuenciaOrden = "Error en N� ORDEN: falta el n�mero " & (penultimoVal + 1)
    End If
    
End Function



Public Function PrimerHuecoFactura(ws As Worksheet, _
                                  ByVal columna As String, _
                                  ByVal filaInicio As Long) As Long
    Dim dict As Scripting.Dictionary
    Dim ultimaFila As Long
    Dim f As Long
    Dim valor As Variant
    Dim n As Long
    
    Set dict = New Scripting.Dictionary
    
    ' �ltima fila con algo en esa columna (aunque sea texto)
    ultimaFila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
    
    ' Si no hay datos por debajo de la fila de inicio ? devolver 1
    If ultimaFila < filaInicio Then
        PrimerHuecoFactura = 1
        Exit Function
    End If
    
    ' Recorrer solo desde filaInicio hasta ultimaFila
    For f = filaInicio To ultimaFila
        valor = ws.Cells(f, columna).Value
        
        ' Limpiar espacios
        If VarType(valor) = vbString Then
            valor = Trim(valor)
        End If
        
        ' Ignorar vac�os, cadenas vac�as, etc.
        If Not IsEmpty(valor) And valor <> "" Then
            ' Intentar convertir a n�mero
            On Error Resume Next
            n = CLng(valor)
            If Err.Number = 0 Then
                If n > 0 Then
                    dict(n) = True
                End If
            End If
            On Error GoTo 0
        End If
    Next f
    
    ' Si no se ha encontrado ning�n n�mero ? devolver 1
    If dict.Count = 0 Then
        PrimerHuecoFactura = 1
        Exit Function
    End If
    
    ' Buscar el primer n�mero libre empezando en 1
    n = 1
    Do While dict.Exists(n)
        n = n + 1
    Loop
    
    PrimerHuecoFactura = n
End Function



' =====================================================================================
' FUNCI�N: PrimerHuecoLibre
' Devuelve el primer n�mero positivo que NO existe en la columna indicada.
' Ejemplo: si existen 1,2,4,5 ? devuelve 3.
' Se usa para asignar n�meros correlativos sin depender del valor m�ximo.
' =====================================================================================
Public Function PrimerHuecoLibreEnColumna(ws As Worksheet, _
                                          ByVal columna As String, _
                                          ByVal filaInicio As Long) As Long
    Dim dict As Scripting.Dictionary
    Dim ultimaFila As Long
    Dim f As Long
    Dim valor As Variant
    Dim i As Long
    
    Set dict = New Scripting.Dictionary
    
    ' �ltima fila con datos en esa columna
    ultimaFila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
    If ultimaFila < filaInicio Then
        ' No hay datos ? el primer n�mero libre es 1
        PrimerHuecoLibreEnColumna = 1
        Exit Function
    End If
    
    ' Recorrer solo desde filaInicio hasta la �ltima fila con datos
    For f = filaInicio To ultimaFila
        valor = ws.Cells(f, columna).Value
        If IsNumeric(valor) Then
            If CLng(valor) > 0 Then
                dict(CLng(valor)) = True
            End If
        End If
    Next f
    
    ' Buscar el primer n�mero libre empezando en 1
    i = 1
    Do While dict.Exists(i)
        i = i + 1
    Loop
    
    PrimerHuecoLibreEnColumna = i
End Function


' =====================================================================================
' FUNCI�N: ObtenerHuecos
' Devuelve una lista de n�meros faltantes entre el m�nimo y el m�ximo del rango.
' Se usa solo para informar al usuario, no para asignar n�meros.
' =====================================================================================
Public Function ObtenerHuecos(rango As Range) As String
    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary

    Dim arrValores As Variant
    Dim i As Long
    Dim valor As Variant
    Dim minVal As Long, maxVal As Long
    Dim firstVal As Boolean
    firstVal = True

    arrValores = rango.Value

    ' Recorrer valores y registrar n�meros
    If rango.Rows.Count = 1 Or rango.Columns.Count > 1 Then
        For i = 1 To rango.Cells.Count
            valor = rango.Cells(i).Value
            If IsNumeric(valor) Then
                dict(valor) = 1
                If firstVal Then
                    minVal = valor: maxVal = valor: firstVal = False
                Else
                    If valor < minVal Then minVal = valor
                    If valor > maxVal Then maxVal = valor
                End If
            End If
        Next i
    ElseIf IsArray(arrValores) Then
        For i = 1 To UBound(arrValores, 1)
            valor = arrValores(i, 1)
            If IsNumeric(valor) Then
                dict(valor) = 1
                If firstVal Then
                    minVal = valor: maxVal = valor: firstVal = False
                Else
                    If valor < minVal Then minVal = valor
                    If valor > maxVal Then maxVal = valor
                End If
            End If
        Next i
    End If

    ' Si no hay n�meros, no hay huecos
    If dict.Count = 0 Then
        ObtenerHuecos = ""
        Exit Function
    End If

    ' ================================
    ' NUEVA L�GICA: limitar rango
    ' ================================
    If (maxVal - minVal) > 10 Then
        ' Rango demasiado grande ? NO mostrar huecos
        ObtenerHuecos = ""
        Exit Function
    End If

    ' Buscar n�meros faltantes SOLO si el rango es peque�o
    Dim faltantes As String
    Dim j As Long
    faltantes = ""

    For j = minVal To maxVal
        If Not dict.Exists(j) Then
            If faltantes = "" Then
                faltantes = CStr(j)
            Else
                faltantes = faltantes & ", " & CStr(j)
            End If
        End If
    Next j

    ObtenerHuecos = faltantes
End Function


' =====================================================================================
' PROCEDIMIENTO: VerificarDuplicados
' Comprueba si el valor ingresado ya existe en la columna.
' Tambi�n informa de huecos num�ricos (solo informativo).
' =====================================================================================
Public Sub VerificarDuplicados(ByVal Target As Range, ByVal rangoColumna As Range, ByVal nombreColumna As String)

    ' Validar que Target es una �nica celda y no est� vac�a
    If Target.Cells.CountLarge > 1 Or IsEmpty(Target.Value) Then Exit Sub

    Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
    Dim arrValores As Variant
    Dim i As Long
    Dim filaTarget As Long
    Dim valorTarget As Variant

    ' Si el rango est� vac�o, salir
    If Application.WorksheetFunction.CountA(rangoColumna) = 0 Then Exit Sub

    arrValores = rangoColumna.Value

    ' Si Target est� vac�o, salir
    If IsEmpty(Target.Value) Then Exit Sub

    ' Construir diccionario con los valores existentes excepto Target
    If rangoColumna.Rows.Count = 1 Or rangoColumna.Columns.Count > 1 Then
        For i = 1 To rangoColumna.Cells.Count
            If rangoColumna.Cells(i).Address <> Target.Address And Not IsEmpty(rangoColumna.Cells(i).Value) Then
                dict(rangoColumna.Cells(i).Value) = 1
            End If
        Next i
        valorTarget = Target.Value
    Else
        filaTarget = Target.Row - rangoColumna.Cells(1, 1).Row + 1
        For i = 1 To UBound(arrValores, 1)
            If i <> filaTarget And Not IsEmpty(arrValores(i, 1)) Then
                dict(arrValores(i, 1)) = 1
            End If
        Next i
        valorTarget = arrValores(filaTarget, 1)
    End If

    ' Verificar duplicado
    If dict.Exists(valorTarget) Then
        MsgBox nombreColumna & " N� " & valorTarget & " YA EXISTE", vbExclamation
        Target.ClearContents
        Target.Select
        Exit Sub
    End If

    ' Informar huecos
    Dim huecos As String
    huecos = ObtenerHuecos(rangoColumna)
    If huecos <> "" Then
        MsgBox nombreColumna & " - N�meros faltantes: " & huecos, vbInformation
        Target.Select
    End If
End Sub

