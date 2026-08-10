Attribute VB_Name = "ModuloLavanderia_Oviedo"
Option Explicit

'==============================================================================
' M�DULO:  ModuloLavanderia_Oviedo
'
' FUNCIONES:
' - Confirmar env�o a lavander�a Oviedo (bot�n CONFIRMAR ENV�O)
' - Registra UNA SOLA FILA con todos los datos TOTALES
' - Excluye autom�ticamente N� ORDEN ya enviados
' - Doble confirmaci�n antes de enviar
'==============================================================================

'---------------------------------------------------------------
' CONFIRMAR ENV�O - Llamada desde el bot�n en Lavander�a Oviedo
'---------------------------------------------------------------
Public Sub ConfirmarEnvioOviedo()

    Dim hoja As Worksheet
    Dim ultimaFilaDatos As Long
    Dim filaTotal As Long
    Dim fila As Long
    Dim numOrden As Variant
    Dim listaOrdenesCompleta As String
    Dim listaOrdenesValidas As String
    Dim ordenesDuplicadas As String
    Dim ordenesExcluidas As String
    Dim respuesta As VbMsgBoxResult
    Dim cantidadOrdenesTotal As Long
    Dim cantidadOrdenesValidas As Long

    Set hoja = ThisWorkbook.Worksheets("Lavander�a Oviedo")

    ' Limpiar hist�rico antiguo (>3 meses)
    Call ModuloLavanderia_Core.LimpiarHistoricoAntiguo("HISTORICO_LAVANDERIA_OVIEDO")

    ' Obtener filas clave
    ultimaFilaDatos = ObtenerUltimaFilaDatos(hoja)
    filaTotal = BuscarFilaTotal(hoja)

    If ultimaFilaDatos < 2 Then
        MsgBox "No hay registros para enviar a lavander�a.", vbInformation
        Exit Sub
    End If

    If filaTotal = 0 Then
        MsgBox "No se encontr� la fila TOTAL.", vbExclamation
        Exit Sub
    End If

    ' Construir lista de N� ORDEN separados por comas
    listaOrdenesCompleta = ""
    cantidadOrdenesTotal = 0
    
    For fila = 2 To ultimaFilaDatos
        numOrden = hoja.Cells(fila, "A").Value
        
        ' Ignorar filas sin N� ORDEN
        If Trim(CStr(numOrden)) = "" Then GoTo SiguienteFila
        
        ' Ignorar fila TOTAL
        If UCase(Trim(hoja.Cells(fila, "B").Value)) = "TOTAL" Then GoTo SiguienteFila
        
        ' A�adir a la lista
        If listaOrdenesCompleta = "" Then
            listaOrdenesCompleta = Trim(CStr(numOrden))
        Else
            listaOrdenesCompleta = listaOrdenesCompleta & ", " & Trim(CStr(numOrden))
        End If
        cantidadOrdenesTotal = cantidadOrdenesTotal + 1
        
SiguienteFila:
    Next fila

    If cantidadOrdenesTotal = 0 Then
        MsgBox "No hay N� ORDEN v�lidos para enviar.", vbInformation
        Exit Sub
    End If

    '----------------------------------------------------------
    ' Verificar duplicados
    '----------------------------------------------------------
    ordenesDuplicadas = ModuloLavanderia_Core.ObtenerOrdenesDuplicadas( _
                            listaOrdenesCompleta, _
                            "HISTORICO_LAVANDERIA_OVIEDO")

    If ordenesDuplicadas <> "" Then
        respuesta = MsgBox( _
            "Los siguientes N� ORDEN ya fueron enviados:" & vbCrLf & vbCrLf & _
            ordenesDuplicadas & vbCrLf & vbCrLf & _
            "Ser�n excluidos del env�o." & vbCrLf & _
            "�Desea continuar?", _
            vbQuestion + vbYesNo)

        If respuesta = vbNo Then Exit Sub

        listaOrdenesValidas = ExcluirOrdenesDuplicadas(listaOrdenesCompleta, ordenesDuplicadas)
        ordenesExcluidas = ordenesDuplicadas
    Else
        listaOrdenesValidas = listaOrdenesCompleta
        ordenesExcluidas = ""
    End If

    If Trim(listaOrdenesValidas) = "" Then
        MsgBox "No quedan N� ORDEN v�lidos para enviar.", vbInformation
        Exit Sub
    End If

    cantidadOrdenesValidas = ContarOrdenes(listaOrdenesValidas)

    '----------------------------------------------------------
    ' Confirmaciones
    '----------------------------------------------------------
    respuesta = MsgBox( _
        "N� ORDEN a enviar: " & listaOrdenesValidas & vbCrLf & _
        "Cantidad: " & cantidadOrdenesValidas & vbCrLf & vbCrLf & _
        "�Desea continuar?", _
        vbQuestion + vbYesNo)

    If respuesta = vbNo Then Exit Sub

    respuesta = MsgBox( _
        "ATENCI�N: El env�o quedar� registrado y no podr� modificarse." & vbCrLf & _
        "�Confirmar env�o?", _
        vbExclamation + vbYesNo)

    If respuesta = vbNo Then Exit Sub

    '----------------------------------------------------------
    ' Solicitar fecha de env�o
    '----------------------------------------------------------
    Dim fechaEnvio As Variant
    Dim fechaValida As Boolean
    Dim fechaTexto As String
    Dim respuestaFecha As VbMsgBoxResult
    Dim diasDiferencia As Long

    fechaValida = False
    Do While Not fechaValida
        fechaTexto = InputBox( _
            "Introduzca la fecha de env�o a lavander�a:" & vbCrLf & vbCrLf & _
            "(Formato: DD/MM/AAAA o DD-MM-AAAA)", _
            "Fecha de env�o", _
            Format(Date, "DD/MM/YYYY"))

        ' Si cancela o deja vac�o, salir
        If Trim(fechaTexto) = "" Then
            MsgBox "Env�o cancelado por el usuario.", vbInformation, "Cancelado"
            Exit Sub
        End If

        '--------------------------------------------------------------
        ' VALIDACI�N 1: Formato correcto (DD/MM/AAAA o DD-MM-AAAA)
        '--------------------------------------------------------------
        If Not ModuloLavanderia_Core.EsFormatoFechaValido(fechaTexto) Then
            MsgBox "Formato incorrecto." & vbCrLf & vbCrLf & _
                   "Por favor, use el formato DD/MM/AAAA o DD-MM-AAAA", _
                   vbExclamation, "Formato no v�lido"
            GoTo SiguienteIntento
        End If

        '--------------------------------------------------------------
        ' VALIDACI�N 2: Fecha v�lida (existe en el calendario)
        '--------------------------------------------------------------
        If Not ModuloLavanderia_Core.EsFechaValida(fechaTexto) Then
            MsgBox "La fecha introducida no existe." & vbCrLf & vbCrLf & _
                   "Por favor, introduzca una fecha v�lida.", _
                   vbExclamation, "Fecha no v�lida"
            GoTo SiguienteIntento
        End If

        ' Convertir a fecha
        fechaEnvio = ModuloLavanderia_Core.ConvertirTextoAFecha(fechaTexto)

        '--------------------------------------------------------------
        ' VALIDACI�N 3: Rango de fechas (-3 / +4 d�as respecto a hoy)
        '--------------------------------------------------------------
        diasDiferencia = DateDiff("d", Date, fechaEnvio)

        If diasDiferencia < -3 Or diasDiferencia > 4 Then
            respuestaFecha = MsgBox( _
                "La fecha introducida (" & Format(fechaEnvio, "DD/MM/YYYY") & ") est� fuera del rango habitual." & vbCrLf & vbCrLf & _
                "�Es correcta la fecha de recogida?", _
                vbQuestion + vbYesNo, _
                "Verificar fecha")

            If respuestaFecha = vbNo Then
                GoTo SiguienteIntento
            End If
        End If

        ' Todas las validaciones pasadas
        fechaValida = True

SiguienteIntento:
    Loop

    '----------------------------------------------------------
    ' Registrar en hist�rico
    '----------------------------------------------------------
    Call ModuloLavanderia_Core.RegistrarEnvioEnHistorico( _
            hoja, _
            filaTotal, _
            listaOrdenesValidas, _
            "HISTORICO_LAVANDERIA_OVIEDO", _
            fechaEnvio)

    MsgBox "Env�o registrado correctamente." & vbCrLf & _
           "N� ORDEN: " & listaOrdenesValidas & vbCrLf & _
           "Fecha de env�o: " & Format(fechaEnvio, "DD/MM/YYYY"), _
           vbInformation

    '----------------------------------------------------------
    ' LIMPIAR LOS DATOS DE LA HOJA LAVANDER�A OVIEDO TRAS EL ENV�O
    '----------------------------------------------------------
    ' Elimina todas las filas de datos salvo la cabecera y la/s fila/s TOTAL
    Dim firstDataRow As Long, lastDataRow As Long
    Dim colFinal As String
    Dim filaBorrar As Long

    colFinal = "U" ' Ajusta si tienes m�s columnas de datos reales
    firstDataRow = 2 ' Cambia si tus datos empiezan en otra fila (ejemplo: 7)

    lastDataRow = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row

    If lastDataRow >= firstDataRow Then
        For filaBorrar = lastDataRow To firstDataRow Step -1
            ' No borrar la fila TOTAL (columna B = "TOTAL")
            If UCase(Trim(hoja.Cells(filaBorrar, "B").Value)) = "TOTAL" Then
                ' Saltar fila de total
            ElseIf Application.WorksheetFunction.CountA(hoja.Range("A" & filaBorrar & ":" & colFinal & filaBorrar)) > 0 Then
                hoja.Range("A" & filaBorrar & ":" & colFinal & filaBorrar).ClearContents
            End If
        Next filaBorrar
    End If

End Sub '---------------------------------------------------------------
' FUNCIONES AUXILIARES (id�nticas a Gij�n)
'---------------------------------------------------------------
Private Function ExcluirOrdenesDuplicadas(ByVal listaCompleta As String, _
                                           ByVal listaDuplicadas As String) As String
    Dim a As Variant, d As Variant
    Dim i As Long, j As Long
    Dim ok As Boolean
    Dim res As String

    a = Split(listaCompleta, ",")
    d = Split(listaDuplicadas, ",")

    For i = LBound(a) To UBound(a)
        ok = True
        For j = LBound(d) To UBound(d)
            If Trim(a(i)) = Trim(d(j)) Then ok = False
        Next j
        If ok Then
            If res = "" Then res = Trim(a(i)) Else res = res & ", " & Trim(a(i))
        End If
    Next i

    ExcluirOrdenesDuplicadas = res
End Function

Private Function ContarOrdenes(ByVal lista As String) As Long
    If Trim(lista) = "" Then Exit Function
    ContarOrdenes = UBound(Split(lista, ",")) + 1
End Function

Private Function ObtenerUltimaFilaDatos(ByVal hoja As Worksheet) As Long
    Dim f As Long
    For f = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If Trim(hoja.Cells(f, "A").Value) <> "" Then
            ObtenerUltimaFilaDatos = f
            Exit Function
        End If
    Next f
End Function

Private Function BuscarFilaTotal(ByVal hoja As Worksheet) As Long
    Dim f As Long
    For f = 2 To hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row
        If UCase(Trim(hoja.Cells(f, "B").Value)) = "TOTAL" Then
            BuscarFilaTotal = f
            Exit Function
        End If
    Next f
End Function


