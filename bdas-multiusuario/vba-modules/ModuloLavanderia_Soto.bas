Attribute VB_Name = "ModuloLavanderia_Soto"
Option Explicit

'==============================================================================
' M�DULO:   ModuloLavanderia_Soto
'
' FUNCIONES:
' - Confirmar env�o a lavander�a (bot�n CONFIRMAR ENV�O)
' - Registra UNA SOLA FILA con todos los datos TOTALES
' - Excluye autom�ticamente N� ORDEN ya enviados
' - Doble confirmaci�n antes de enviar
'==============================================================================

'---------------------------------------------------------------
' CONFIRMAR ENV�O - Llamada desde el bot�n en Lavander�a Soto
'---------------------------------------------------------------
Public Sub ConfirmarEnvioSoto()
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
    Dim ordenes As Variant
    Dim i As Long
    
    Set hoja = ThisWorkbook.Worksheets("Lavander�a Soto")
    
    ' Primero, limpiar registros antiguos del hist�rico (m�s de 3 meses)
    Call ModuloLavanderia_Core.LimpiarHistoricoAntiguo("HISTORICO_LAVANDERIA_SOTO")
    
    ' Obtener �ltima fila con datos y fila TOTAL
    ultimaFilaDatos = ObtenerUltimaFilaDatos(hoja)
    filaTotal = BuscarFilaTotal(hoja)
    
    ' Verificar que hay datos para enviar
    If ultimaFilaDatos < 2 Then
        MsgBox "No hay registros para enviar a lavander�a.", vbInformation, "Sin datos"
        Exit Sub
    End If
    
    ' Verificar que existe la fila TOTAL
    If filaTotal = 0 Then
        MsgBox "No se encontr� la fila TOTAL." & vbCrLf & _
               "Verifique que existan datos v�lidos.", vbExclamation, "Error"
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
    
    ' Verificar que hay �rdenes para enviar
    If cantidadOrdenesTotal = 0 Then
        MsgBox "No hay N� ORDEN v�lidos para enviar.", vbInformation, "Sin datos"
        Exit Sub
    End If
    
    '==========================================================
    ' VERIFICAR DUPLICADOS Y EXCLUIRLOS
    '==========================================================
    ordenesDuplicadas = ModuloLavanderia_Core.ObtenerOrdenesDuplicadas(listaOrdenesCompleta, "HISTORICO_LAVANDERIA_SOTO")
    
    If ordenesDuplicadas <> "" Then
        ' Notificar duplicados encontrados
        respuesta = MsgBox( _
            "Los siguientes N� ORDEN ya fueron enviados a lavander�a anteriormente:" & vbCrLf & vbCrLf & _
            ordenesDuplicadas & vbCrLf & vbCrLf & _
            "Estos N� ORDEN ser�n EXCLUIDOS del env�o actual." & vbCrLf & vbCrLf & _
            "�Desea continuar con el env�o de los N� ORDEN restantes?", _
            vbQuestion + vbYesNo, _
            "�rdenes duplicadas detectadas")
        
        If respuesta = vbNo Then Exit Sub
        
        ' Construir lista excluyendo duplicados
        listaOrdenesValidas = ExcluirOrdenesDuplicadas(listaOrdenesCompleta, ordenesDuplicadas)
        ordenesExcluidas = ordenesDuplicadas
    Else
        listaOrdenesValidas = listaOrdenesCompleta
        ordenesExcluidas = ""
    End If
    
    ' Contar �rdenes v�lidas
    If Trim(listaOrdenesValidas) = "" Then
        MsgBox "No quedan N� ORDEN v�lidos para enviar." & vbCrLf & _
               "Todos los N� ORDEN ya fueron enviados anteriormente.", _
               vbInformation, "Sin datos nuevos"
        Exit Sub
    End If
    
    cantidadOrdenesValidas = ContarOrdenes(listaOrdenesValidas)
    
    '==========================================================
    ' PRIMERA CONFIRMACI�N:  Resumen del env�o
    '==========================================================
    Dim mensajeConfirmacion As String
    mensajeConfirmacion = "Se va a registrar el env�o con los siguientes datos:" & vbCrLf & vbCrLf & _
                          "� N� ORDEN a enviar: " & listaOrdenesValidas & vbCrLf & _
                          "� Cantidad de pedidos: " & cantidadOrdenesValidas
    
    If ordenesExcluidas <> "" Then
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "� N� ORDEN EXCLUIDOS (ya enviados): " & ordenesExcluidas
    End If
    
    mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & "�Desea continuar?"
    
    respuesta = MsgBox(mensajeConfirmacion, vbQuestion + vbYesNo, "Confirmar env�o - Paso 1 de 2")
    
    If respuesta = vbNo Then Exit Sub
    
    '==========================================================
    ' SEGUNDA CONFIRMACI�N: Aviso de irreversibilidad
    '==========================================================
    respuesta = MsgBox( _
        "�ATENCI�N!" & vbCrLf & vbCrLf & _
        "Una vez confirmado el env�o:" & vbCrLf & vbCrLf & _
        "� NO ser� posible modificar la selecci�n de prendas." & vbCrLf & _
        "� Los datos quedar�n registrados en el hist�rico." & vbCrLf & _
        "� Los N� ORDEN enviados no podr�n volver a enviarse." & vbCrLf & vbCrLf & _
        "�Est� SEGURO de que desea confirmar el env�o?", _
        vbExclamation + vbYesNo, _
        "Confirmar env�o - Paso 2 de 2")
    
    If respuesta = vbNo Then Exit Sub
    
    '==========================================================
    ' SOLICITAR FECHA DE ENV�O
    '==========================================================
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
        If Not EsFormatoFechaValido(fechaTexto) Then
            MsgBox "Formato incorrecto." & vbCrLf & vbCrLf & _
                   "Por favor, use el formato DD/MM/AAAA o DD-MM-AAAA", _
                   vbExclamation, "Formato no v�lido"
            GoTo SiguienteIntento
        End If
        
        '--------------------------------------------------------------
        ' VALIDACI�N 2: Fecha v�lida (existe en el calendario)
        '--------------------------------------------------------------
        If Not EsFechaValida(fechaTexto) Then
            MsgBox "La fecha introducida no existe." & vbCrLf & vbCrLf & _
                   "Por favor, introduzca una fecha v�lida.", _
                   vbExclamation, "Fecha no v�lida"
            GoTo SiguienteIntento
        End If
        
        ' Convertir a fecha
        fechaEnvio = ConvertirTextoAFecha(fechaTexto)
        
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
    
    '==========================================================
    ' REGISTRAR EN HIST�RICO
    '==========================================================
    Call ModuloLavanderia_Core.RegistrarEnvioEnHistorico(hoja, filaTotal, listaOrdenesValidas, "HISTORICO_LAVANDERIA_SOTO", fechaEnvio)
    
    '==========================================================
    ' MOSTRAR RESUMEN FINAL
    '==========================================================
    Dim mensajeFinal As String
    mensajeFinal = "Env�o registrado correctamente." & vbCrLf & vbCrLf & _
                   "� N� ORDEN enviados: " & listaOrdenesValidas & vbCrLf & _
                   "� Fecha de env�o: " & Format(Date, "DD/MM/YYYY") & vbCrLf & _
                   "� Usuario: " & Trim(ThisWorkbook.usuarioActual)
    
    If ordenesExcluidas <> "" Then
        mensajeFinal = mensajeFinal & vbCrLf & vbCrLf & _
                       "� N� ORDEN EXCLUIDOS: " & ordenesExcluidas
    End If
    
    MsgBox mensajeFinal, vbInformation, "Env�o completado"
    
    '==========================================================
    ' LIMPIAR LOS DATOS EN LA HOJA LAVANDER�A SOTO TRAS EL ENV�O
    ' Elimina todas las filas de datos salvo la cabecera y la/s fila/s TOTAL
    '==========================================================
    Dim firstDataRow As Long, lastDataRow As Long
    Dim colFinal As String
    Dim filaBorrar As Long

    colFinal = "U" ' Ajusta si tienes m�s columnas reales de datos
    firstDataRow = 2 ' Cambia si tu tabla de datos empieza en otra fila (ejemplo: 7)

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

End Sub
    

'---------------------------------------------------------------
' Excluye los N� ORDEN duplicados de la lista original
'---------------------------------------------------------------
Private Function ExcluirOrdenesDuplicadas(ByVal listaCompleta As String, _
                                           ByVal listaDuplicadas As String) As String
    Dim ordenesCompletas As Variant
    Dim ordenesDuplicadas As Variant
    Dim resultado As String
    Dim i As Long, j As Long
    Dim esDuplicada As Boolean
    
    ordenesCompletas = Split(listaCompleta, ",")
    ordenesDuplicadas = Split(listaDuplicadas, ",")
    resultado = ""
    
    For i = LBound(ordenesCompletas) To UBound(ordenesCompletas)
        esDuplicada = False
        
        ' Verificar si est� en la lista de duplicadas
        For j = LBound(ordenesDuplicadas) To UBound(ordenesDuplicadas)
            If Trim(ordenesCompletas(i)) = Trim(ordenesDuplicadas(j)) Then
                esDuplicada = True
                Exit For
            End If
        Next j
        
        ' Si no es duplicada, a�adir al resultado
        If Not esDuplicada Then
            If resultado = "" Then
                resultado = Trim(ordenesCompletas(i))
            Else
                resultado = resultado & ", " & Trim(ordenesCompletas(i))
            End If
        End If
    Next i
    
    ExcluirOrdenesDuplicadas = resultado
End Function

'---------------------------------------------------------------
' Cuenta la cantidad de N� ORDEN en una lista separada por comas
'---------------------------------------------------------------
Private Function ContarOrdenes(ByVal listaOrdenes As String) As Long
    Dim ordenes As Variant
    
    If Trim(listaOrdenes) = "" Then
        ContarOrdenes = 0
        Exit Function
    End If
    
    ordenes = Split(listaOrdenes, ",")
    ContarOrdenes = UBound(ordenes) - LBound(ordenes) + 1
End Function

'---------------------------------------------------------------
' Obtener �ltima fila con datos (excluyendo TOTAL)
'---------------------------------------------------------------
Private Function ObtenerUltimaFilaDatos(ByVal hoja As Worksheet) As Long
    Dim fila As Long
    Dim ultimaFila As Long
    
    ultimaFila = Application.WorksheetFunction.Max( _
        hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row)
    
    For fila = ultimaFila To 2 Step -1
        ' Ignorar filas de TOTAL
        If UCase(Trim(hoja.Cells(fila, "B").Value)) = "TOTAL" Then
            GoTo SiguienteFila
        End If
        
        ' Ignorar filas vac�as
        If Application.WorksheetFunction.CountA(hoja.Range("A" & fila & ":U" & fila)) = 0 Then
            GoTo SiguienteFila
        End If
        
        ' Fila v�lida con datos
        If Trim(CStr(hoja.Cells(fila, "A").Value)) <> "" Then
            ObtenerUltimaFilaDatos = fila
            Exit Function
        End If
        
SiguienteFila:
    Next fila
    
    ObtenerUltimaFilaDatos = 1
End Function

'---------------------------------------------------------------
' Buscar fila donde est� "TOTAL"
'---------------------------------------------------------------
Private Function BuscarFilaTotal(ByVal hoja As Worksheet) As Long
    Dim fila As Long
    Dim ultimaBusqueda As Long
    
    ultimaBusqueda = hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row + 10
    
    For fila = 2 To ultimaBusqueda
        If UCase(Trim(hoja.Cells(fila, "B").Value)) = "TOTAL" Then
            BuscarFilaTotal = fila
            Exit Function
        End If
    Next fila
    
    BuscarFilaTotal = 0
End Function

