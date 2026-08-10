Attribute VB_Name = "CopiarDatosFacturaEstudiante"
Sub TransferirDatosFactura()
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual
    Application.enableEvents = False
    
    On Error GoTo CleanUp
    
    Dim wsOrigen As Worksheet
    Dim wsDestino As Worksheet
    Dim numero As Long
    Dim ano As Long
    Dim filaEncontrada As Long
    Dim encontrado As Boolean
    
    ' Asignar hojas
    Set wsOrigen = ThisWorkbook.Sheets("RESIDENCIA ESTUDIANTES")
    Set wsDestino = ThisWorkbook.Sheets("FACTURA R. ESTUDIANTES")
    
    ' Validar datos de entrada
    If Not IsNumeric(wsDestino.Range("C1").Value) Or _
       Not IsNumeric(wsDestino.Range("D1").Value) Then
        MsgBox "Los valores en C1 y D1 deben ser num�ricos", vbExclamation
        Exit Sub
    End If
    
    numero = wsDestino.Range("C1").Value
    ano = wsDestino.Range("D1").Value
    
    ' Buscar coincidencia usando Find
    With wsOrigen.Columns("A")
        Dim celda As Range
        Set celda = .Find(What:=numero, _
                          LookIn:=xlValues, _
                          LookAt:=xlWhole)
        
        If Not celda Is Nothing Then
            Dim primeraDireccion As String
            primeraDireccion = celda.Address
            
            Do
                If wsOrigen.Cells(celda.Row, "B").Value = ano Then
                    filaEncontrada = celda.Row
                    encontrado = True
                    Exit Do
                End If
                Set celda = .FindNext(celda)
            Loop While Not celda Is Nothing And celda.Address <> primeraDireccion
        End If
    End With
    
    If Not encontrado Then
        MsgBox "No se encontr� la combinaci�n N�mero/A�o especificada", vbExclamation
        Exit Sub
    End If
    
    ' Transferir datos usando arrays para mejor rendimiento
    With wsOrigen
        Dim datosOrigen As Variant
        datosOrigen = .Rows(filaEncontrada).Value
    End With
    
    With wsDestino
        ' Mapeo de columnas a celdas destino
        .Range("F14").Value = datosOrigen(1, 6)  ' Col F - F14=FECHA EMISI�N
        .Range("R36").Value = datosOrigen(1, 10) ' Col J - R36=%FAM.NUM.
        .Range("T23").Value = datosOrigen(1, 9)  ' Col I - T23=NETO
        .Range("R35").Value = datosOrigen(1, 11) ' Col K - R35=%INCREMENTO
        .Range("R37").Value = datosOrigen(1, 12) ' Col L - R37=%IVA
        .Range("T38").Value = datosOrigen(1, 13) ' Col M - T38=TOTAL FACTURA
        .Range("F15").Value = datosOrigen(1, 4)  ' Col D - F15=NOMBRE
        .Range("B23").Value = datosOrigen(1, 7)  ' Col G - B23=DIAS ESTANCIA
        .Range("F19").Value = datosOrigen(1, 5)  ' Col E - F19=FECHA ENTRADA
        .Range("F20").Value = datosOrigen(1, 6)  ' Col F - F20=FECHA SALIDA
        .Range("F17").Value = datosOrigen(1, 8)  ' Col H - F17=DIRECCI�N
        .Range("F16").Value = datosOrigen(1, 3)  ' Col C - F16=DNI
        .Range("E23").Value = datosOrigen(1, 14) ' Col N - E23=N� HABITACI�N
        .Range("T20").Value = datosOrigen(1, 2)  ' Col B - T20=A�O FACTURA
        .Range("R20").Value = datosOrigen(1, 1)  ' Col A - R20=N� FACTURA
        
        ' NUEVOS C�LCULOS (sustituyendo las f�rmulas)
        Call CalcularFormulas(wsDestino)
    End With
    
CleanUp:
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
End Sub

Sub CalcularFormulas(ws As Worksheet)
    Dim T34 As Double, T35 As Double, T36 As Double, T37 As Double
    Dim T23 As Double, F20 As Date, F19 As Date
    Dim R35 As Double, R36 As Double, R37 As Double
    Dim diasEnMes As Integer
    
    ' Obtener valores necesarios
    T23 = ws.Range("T23").Value
    F20 = ws.Range("F20").Value
    F19 = ws.Range("F19").Value
    R35 = ws.Range("R35").Value
    R36 = ws.Range("R36").Value
    R37 = ws.Range("R37").Value
    
    ' Calcular d�as del mes de F20
    diasEnMes = Day(DateSerial(Year(F20), Month(F20) + 1, 0))
    
    ' T34: =T23*((F20-F19)+1)/DIA(F20)
    T34 = T23 * ((F20 - F19) + 1) / diasEnMes
    ws.Range("T34").Value = T34
    
    ' T35: =T34*R35
    T35 = T34 * R35
    ws.Range("T35").Value = T35
    
    ' T36: =-(T34+T34*R35)*R36
    T36 = -(T34 + T34 * R35) * R36
    ws.Range("T36").Value = T36
    
    ' T37: =SUMA(T34:T36)*R37
    T37 = (T34 + T35 + T36) * R37
    ws.Range("T37").Value = T37
End Sub

