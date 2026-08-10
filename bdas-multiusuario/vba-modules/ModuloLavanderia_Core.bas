Attribute VB_Name = "ModuloLavanderia_Core"
Option Explicit

'==============================================================================
' M�DULO:   ModuloLavanderia_Core
'
' FUNCIONES:
' - Verificar si N� ORDEN existe en hist�rico (busca dentro de cadenas con comas)
' - Registrar env�o en hist�rico (una fila con todos los datos TOTALES)
' - Limpiar registros antiguos (m�s de 3 meses)
'==============================================================================

'---------------------------------------------------------------
' Normaliza texto de alojamientos:
' - Elimina espacios raros
' - Unifica separadores
' - Convierte a MAY�SCULAS
'---------------------------------------------------------------
Public Function NormalizarTextoAlojamientos(ByVal texto As String) As String
    Dim t As String

    t = CStr(texto)

    ' Quitar saltos de l�nea ocultos
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "")

    ' Quitar espacios no separables (MUY IMPORTANTE)
    t = Replace(t, Chr(160), "")

    ' Quitar espacios normales
    t = Replace(t, " ", "")

    ' Pasar a may�sculas
    t = UCase(t)

    NormalizarTextoAlojamientos = t
End Function


Public Function ListaAlojamientosValida( _
        ByVal texto As String, _
        ByVal listaValidos As Variant) As Boolean

    Dim partes As Variant
    Dim i As Long, j As Long
    Dim valor As String
    Dim textoNorm As String
    Dim validoNorm As String
    Dim encontrado As Boolean

    If Trim(texto) = "" Then Exit Function

    ' NORMALIZAR TEXTO COMPLETO
    textoNorm = NormalizarTextoAlojamientos(texto)

    ' Separar por coma (ya sin espacios)
    partes = Split(textoNorm, ",")

    For i = LBound(partes) To UBound(partes)

        valor = partes(i)
        encontrado = False

        For j = LBound(listaValidos) To UBound(listaValidos)
            validoNorm = NormalizarTextoAlojamientos(listaValidos(j))

            If valor = validoNorm Then
                encontrado = True
                Exit For
            End If
        Next j

        If Not encontrado Then
            ListaAlojamientosValida = False
            Exit Function
        End If
    Next i

    ListaAlojamientosValida = True
End Function



'---------------------------------------------------------------
' Comprueba si un N� ORDEN ya existe en el hist�rico
' Busca dentro de la columna C que contiene N� ORDEN separados por comas
'---------------------------------------------------------------
Public Function OrdenYaEnviada(ByVal numOrden As Variant, _
                               ByVal nombreHojaHistorico As String) As Boolean
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim contenidoCelda As String
    Dim ordenes As Variant
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHojaHistorico)
    On Error GoTo 0

    If ws Is Nothing Then
        OrdenYaEnviada = False
        Exit Function
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Recorrer cada fila del hist�rico
    For fila = 2 To ultimaFila
        contenidoCelda = Trim(CStr(ws.Cells(fila, "C").Value))
        
        ' Dividir por comas y buscar coincidencia exacta
        ordenes = Split(contenidoCelda, ",")
        For i = LBound(ordenes) To UBound(ordenes)
            If Trim(ordenes(i)) = Trim(CStr(numOrden)) Then
                OrdenYaEnviada = True
                Exit Function
            End If
        Next i
    Next fila

    OrdenYaEnviada = False
End Function

'---------------------------------------------------------------
' Busca la fila donde est�n EXACTAMENTE los mismos N� ORDEN
' (para sobrescribir si se reenv�a el mismo conjunto)
' Devuelve 0 si no encuentra coincidencia exacta
'---------------------------------------------------------------
Public Function BuscarFilaOrdenesEnHistorico(ByVal listaOrdenes As String, _
                                              ByVal nombreHojaHistorico As String) As Long
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim contenidoCelda As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHojaHistorico)
    On Error GoTo 0

    If ws Is Nothing Then
        BuscarFilaOrdenesEnHistorico = 0
        Exit Function
    End If

    ultimaFila = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row
    
    ' Normalizar la lista para comparaci�n (quitar espacios extra)
    listaOrdenes = NormalizarListaOrdenes(listaOrdenes)
    
    ' Buscar coincidencia exacta
    For fila = 2 To ultimaFila
        contenidoCelda = NormalizarListaOrdenes(CStr(ws.Cells(fila, "C").Value))
        
        If contenidoCelda = listaOrdenes Then
            BuscarFilaOrdenesEnHistorico = fila
            Exit Function
        End If
    Next fila

    BuscarFilaOrdenesEnHistorico = 0
End Function

'---------------------------------------------------------------
' Normaliza una lista de N� ORDEN (quita espacios, ordena)
'---------------------------------------------------------------
Private Function NormalizarListaOrdenes(ByVal lista As String) As String
    Dim ordenes As Variant
    Dim i As Long
    Dim resultado As String
    
    ordenes = Split(lista, ",")
    
    ' Limpiar espacios de cada elemento
    For i = LBound(ordenes) To UBound(ordenes)
        ordenes(i) = Trim(ordenes(i))
    Next i
    
    ' Ordenar para comparaci�n consistente
    Call OrdenarArray(ordenes)
    
    ' Reconstruir cadena
    resultado = Join(ordenes, ", ")
    
    NormalizarListaOrdenes = resultado
End Function

'---------------------------------------------------------------
' Ordena un array de strings (bubble sort simple)
'---------------------------------------------------------------
Private Sub OrdenarArray(ByRef arr As Variant)
    Dim i As Long, j As Long
    Dim temp As String
    
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub

'---------------------------------------------------------------
' Registra un env�o completo en el hist�rico (una sola fila)
' OPCI�N B: Acumula en la misma fila si ya existe env�o de HOY
'
' L�gica:
' - Si existe fila con FECHA ENV�O = hoy ? a�ade N� ORDEN y suma cantidades
' - Si no existe ? crea nueva fila
'---------------------------------------------------------------
Public Sub RegistrarEnvioEnHistorico(ByVal hojaOrigen As Worksheet, _
                                      ByVal filaTotal As Long, _
                                      ByVal listaOrdenes As String, _
                                      ByVal nombreHojaHistorico As String, _
                                      Optional ByVal fechaEnvio As Variant = Null)
    Dim wsHistorico As Worksheet
    Dim filaDestino As Long
    Dim filaExistente As Long
    Dim estabaProtegida As Boolean
    Dim eventosEstaban As Boolean
    Dim pantallaEstaba As Boolean
    Dim ordenesExistentes As String
    Dim ordenesNuevas As String
    Dim ordenesCombinadas As String
    
    On Error Resume Next
    Set wsHistorico = ThisWorkbook.Worksheets(nombreHojaHistorico)
    On Error GoTo 0
    
    If wsHistorico Is Nothing Then
        MsgBox "No se encontr� la hoja '" & nombreHojaHistorico & "'.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Guardar estado actual de eventos y pantalla
    eventosEstaban = Application.enableEvents
    pantallaEstaba = Application.screenUpdating
    
    ' Deshabilitar eventos y actualizaci�n de pantalla ANTES de desproteger
    Application.enableEvents = False
    Application.screenUpdating = False
    
    On Error GoTo ErrorEscritura
    
    ' Verificar si la hoja est� protegida y desprotegerla
    estabaProtegida = wsHistorico.ProtectContents
    If estabaProtegida Then
        On Error Resume Next
        wsHistorico.Unprotect password:=""
        If wsHistorico.ProtectContents Then
            wsHistorico.Unprotect password:=ModuloConfigSegura.ObtenerPasswordLOG()
        End If
        On Error GoTo ErrorEscritura
    End If
    
    ' Si no se proporcion� fecha, usar hoy
    Dim fechaAUsar As Date
    If IsNull(fechaEnvio) Or IsEmpty(fechaEnvio) Then
        fechaAUsar = Date
    Else
        fechaAUsar = CDate(fechaEnvio)
    End If
    
    ' Buscar si existe una fila con la misma FECHA ENV�O
    filaExistente = BuscarFilaEnvioFecha(wsHistorico, fechaAUsar)
    
    If filaExistente > 0 Then
        '==============================================================
        ' EXISTE ENV�O DE HOY ? ACUMULAR
        '==============================================================
        filaDestino = filaExistente
        
        ' Obtener �rdenes existentes y combinar con las nuevas
        ordenesExistentes = Trim(CStr(wsHistorico.Cells(filaDestino, "C").Value))
        ordenesNuevas = listaOrdenes
        ordenesCombinadas = CombinarOrdenes(ordenesExistentes, ordenesNuevas)
        
        ' Actualizar datos acumulando
        With wsHistorico
            ' Usuario:  mantener el existente (no cambiar)
            ' Fecha env�o: mantener la existente (no cambiar)
            
            ' N� ORDEN:  combinar
            .Cells(filaDestino, "C").Value = ordenesCombinadas
            
            ' FECHA DESDE: la menor entre existente y nueva
            .Cells(filaDestino, "D").Value = FechaMenor( _
                .Cells(filaDestino, "D").Value, _
                hojaOrigen.Cells(filaTotal, "C").Value)
            
            ' FECHA HASTA:  la mayor entre existente y nueva
            .Cells(filaDestino, "E").Value = FechaMayor( _
                .Cells(filaDestino, "E").Value, _
                hojaOrigen.Cells(filaTotal, "D").Value)
            
            ' Columnas F a T:  SUMAR los valores nuevos a los existentes
            .Cells(filaDestino, "F").Value = SumarValores(.Cells(filaDestino, "F").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "E"))
            .Cells(filaDestino, "G").Value = SumarValores(.Cells(filaDestino, "G").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "F"))
            .Cells(filaDestino, "H").Value = SumarValores(.Cells(filaDestino, "H").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "G"))
            .Cells(filaDestino, "I").Value = SumarValores(.Cells(filaDestino, "I").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "H"))
            .Cells(filaDestino, "J").Value = SumarValores(.Cells(filaDestino, "J").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "I"))
            .Cells(filaDestino, "K").Value = SumarValores(.Cells(filaDestino, "K").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "J"))
            .Cells(filaDestino, "L").Value = SumarValores(.Cells(filaDestino, "L").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "K"))
            .Cells(filaDestino, "M").Value = SumarValores(.Cells(filaDestino, "M").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "L"))
            .Cells(filaDestino, "N").Value = SumarValores(.Cells(filaDestino, "N").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "O"))
            .Cells(filaDestino, "O").Value = SumarValores(.Cells(filaDestino, "O").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "M"))
            .Cells(filaDestino, "P").Value = SumarValores(.Cells(filaDestino, "P").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "P"))
            .Cells(filaDestino, "Q").Value = SumarValores(.Cells(filaDestino, "Q").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "N"))
            .Cells(filaDestino, "R").Value = SumarValores(.Cells(filaDestino, "R").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "Q"))
            .Cells(filaDestino, "S").Value = SumarValores(.Cells(filaDestino, "S").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "R"))
            .Cells(filaDestino, "T").Value = SumarValores(.Cells(filaDestino, "T").Value, ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "S"))
        End With
        
    Else
        '==============================================================
        ' NO EXISTE ENV�O DE HOY ? CREAR NUEVA FILA
        '==============================================================
        filaDestino = wsHistorico.Cells(wsHistorico.Rows.Count, "A").End(xlUp).Row + 1
        If filaDestino < 2 Then filaDestino = 2
        
        ' Copiar datos al hist�rico (solo de las �rdenes v�lidas)
        With wsHistorico
            .Cells(filaDestino, "A").Value = Trim(ThisWorkbook.usuarioActual)
            .Cells(filaDestino, "B").Value = fechaAUsar
            .Cells(filaDestino, "C").Value = listaOrdenes
            
            ' Obtener fechas solo de las �rdenes v�lidas
            .Cells(filaDestino, "D").Value = ObtenerFechaMinOrdenesValidas(hojaOrigen, listaOrdenes, "C")
            .Cells(filaDestino, "E").Value = ObtenerFechaMaxOrdenesValidas(hojaOrigen, listaOrdenes, "D")
            
            ' Obtener sumas solo de las �rdenes v�lidas
            .Cells(filaDestino, "F").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "E")
            .Cells(filaDestino, "G").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "F")
            .Cells(filaDestino, "H").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "G")
            .Cells(filaDestino, "I").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "H")
            .Cells(filaDestino, "J").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "I")
            .Cells(filaDestino, "K").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "J")
            .Cells(filaDestino, "L").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "K")
            .Cells(filaDestino, "M").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "L")
            .Cells(filaDestino, "N").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "O")
            .Cells(filaDestino, "O").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "M")
            .Cells(filaDestino, "P").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "P")
            .Cells(filaDestino, "Q").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "N")
            .Cells(filaDestino, "R").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "Q")
            .Cells(filaDestino, "S").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "R")
            .Cells(filaDestino, "T").Value = ObtenerValorOrdenesValidas(hojaOrigen, listaOrdenes, "S")
        End With
    End If
    
    ' Formato de fechas y alineaci�n
    With wsHistorico
        .Cells(filaDestino, "B").NumberFormat = "DD/MM/YYYY"
        .Cells(filaDestino, "D").NumberFormat = "DD/MM/YYYY"
        .Cells(filaDestino, "E").NumberFormat = "DD/MM/YYYY"
        .Range(.Cells(filaDestino, "A"), .Cells(filaDestino, "T")).HorizontalAlignment = xlCenter
    End With
    
    ' Reproteger la hoja si estaba protegida
    If estabaProtegida Then
        On Error Resume Next
        wsHistorico.Protect password:=ModuloConfigSegura.ObtenerPasswordLOG(), UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    
    ' Restaurar estado de eventos y pantalla
    Application.screenUpdating = pantallaEstaba
    Application.enableEvents = eventosEstaban
    
    Exit Sub
    
ErrorEscritura:
    If estabaProtegida Then
        On Error Resume Next
        wsHistorico.Protect password:=ModuloConfigSegura.ObtenerPasswordLOG(), UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    
    Application.screenUpdating = pantallaEstaba
    Application.enableEvents = eventosEstaban
    
    MsgBox "Error al escribir en el hist�rico: " & Err.Description, vbCritical, "Error"
End Sub

'---------------------------------------------------------------
' Verifica si el texto tiene formato DD/MM/AAAA o DD-MM-AAAA
'---------------------------------------------------------------
Public Function EsFormatoFechaValido(ByVal textoFecha As String) As Boolean
    Dim partes As Variant
    Dim separador As String
    
    textoFecha = Trim(textoFecha)
    
    ' Determinar separador
    If InStr(textoFecha, "/") > 0 Then
        separador = "/"
    ElseIf InStr(textoFecha, "-") > 0 Then
        separador = "-"
    Else
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    ' Dividir por separador
    partes = Split(textoFecha, separador)
    
    ' Debe tener 3 partes (d�a, mes, a�o)
    If UBound(partes) - LBound(partes) + 1 <> 3 Then
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    ' Verificar que cada parte sea num�rica
    If Not IsNumeric(partes(0)) Or Not IsNumeric(partes(1)) Or Not IsNumeric(partes(2)) Then
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    ' Verificar longitudes (DD = 1-2 d�gitos, MM = 1-2 d�gitos, AAAA = 4 d�gitos)
    If Len(partes(0)) < 1 Or Len(partes(0)) > 2 Then
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    If Len(partes(1)) < 1 Or Len(partes(1)) > 2 Then
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    If Len(partes(2)) <> 4 Then
        EsFormatoFechaValido = False
        Exit Function
    End If
    
    EsFormatoFechaValido = True
End Function

'---------------------------------------------------------------
' Verifica si la fecha existe en el calendario
'---------------------------------------------------------------
Public Function EsFechaValida(ByVal textoFecha As String) As Boolean
    Dim partes As Variant
    Dim separador As String
    Dim dia As Integer
    Dim mes As Integer
    Dim anio As Integer
    Dim fechaPrueba As Date
    
    ' Determinar separador
    If InStr(textoFecha, "/") > 0 Then
        separador = "/"
    Else
        separador = "-"
    End If
    
    partes = Split(textoFecha, separador)
    
    dia = CInt(partes(0))
    mes = CInt(partes(1))
    anio = CInt(partes(2))
    
    ' Validar rangos b�sicos
    If mes < 1 Or mes > 12 Then
        EsFechaValida = False
        Exit Function
    End If
    
    If dia < 1 Or dia > 31 Then
        EsFechaValida = False
        Exit Function
    End If
    
    If anio < 1900 Or anio > 2100 Then
        EsFechaValida = False
        Exit Function
    End If
    
    ' Intentar crear la fecha (esto valida d�as por mes, a�os bisiestos, etc.)
    On Error GoTo FechaInvalida
    fechaPrueba = DateSerial(anio, mes, dia)
    
    ' Verificar que la fecha creada corresponde a los valores introducidos
    ' (DateSerial puede ajustar fechas inv�lidas, ej: 31/02 -> 03/03)
    If Day(fechaPrueba) <> dia Or Month(fechaPrueba) <> mes Or Year(fechaPrueba) <> anio Then
        EsFechaValida = False
        Exit Function
    End If
    
    EsFechaValida = True
    Exit Function
    
FechaInvalida:
    EsFechaValida = False
End Function

'---------------------------------------------------------------
' Convierte texto DD/MM/AAAA o DD-MM-AAAA a fecha
'---------------------------------------------------------------
Public Function ConvertirTextoAFecha(ByVal textoFecha As String) As Date
    Dim partes As Variant
    Dim separador As String
    Dim dia As Integer
    Dim mes As Integer
    Dim anio As Integer
    
    ' Determinar separador
    If InStr(textoFecha, "/") > 0 Then
        separador = "/"
    Else
        separador = "-"
    End If
    
    partes = Split(textoFecha, separador)
    
    dia = CInt(partes(0))
    mes = CInt(partes(1))
    anio = CInt(partes(2))
    
    ConvertirTextoAFecha = DateSerial(anio, mes, dia)
End Function

'---------------------------------------------------------------
' Busca si existe una fila con FECHA ENV�O = HOY
' Devuelve el n�mero de fila o 0 si no existe
'---------------------------------------------------------------
Private Function BuscarFilaEnvioFecha(ByVal wsHistorico As Worksheet, ByVal fechaBuscar As Date) As Long
    Dim ultimaFila As Long
    Dim fila As Long
    Dim fechaEnvio As Date
    
    ultimaFila = wsHistorico.Cells(wsHistorico.Rows.Count, "B").End(xlUp).Row
    
    For fila = 2 To ultimaFila
        If IsDate(wsHistorico.Cells(fila, "B").Value) Then
            fechaEnvio = CDate(wsHistorico.Cells(fila, "B").Value)
            If fechaEnvio = fechaBuscar Then
                BuscarFilaEnvioFecha = fila
                Exit Function
            End If
        End If
    Next fila
    
    BuscarFilaEnvioFecha = 0
End Function

'---------------------------------------------------------------
' Combina dos listas de �rdenes eliminando duplicados
'---------------------------------------------------------------
Private Function CombinarOrdenes(ByVal ordenesExistentes As String, ByVal ordenesNuevas As String) As String
    Dim todasOrdenes As String
    Dim arrExistentes As Variant
    Dim arrNuevas As Variant
    Dim resultado As String
    Dim i As Long
    Dim orden As String
    
    resultado = ordenesExistentes
    
    If Trim(ordenesNuevas) = "" Then
        CombinarOrdenes = resultado
        Exit Function
    End If
    
    arrNuevas = Split(ordenesNuevas, ",")
    
    For i = LBound(arrNuevas) To UBound(arrNuevas)
        orden = Trim(arrNuevas(i))
        ' Solo a�adir si no existe ya
        If InStr("," & Replace(resultado, " ", "") & ",", "," & orden & ",") = 0 Then
            If resultado = "" Then
                resultado = orden
            Else
                resultado = resultado & ", " & orden
            End If
        End If
    Next i
    
    CombinarOrdenes = resultado
End Function

'---------------------------------------------------------------
' Suma dos valores (maneja vac�os y no num�ricos)
'---------------------------------------------------------------
Private Function SumarValores(ByVal valor1 As Variant, ByVal valor2 As Variant) As Variant
    Dim v1 As Double
    Dim v2 As Double
    
    v1 = 0
    v2 = 0
    
    If IsNumeric(valor1) And Not IsEmpty(valor1) Then v1 = CDbl(valor1)
    If IsNumeric(valor2) And Not IsEmpty(valor2) Then v2 = CDbl(valor2)
    
    If v1 + v2 = 0 Then
        SumarValores = ""
    Else
        SumarValores = v1 + v2
    End If
End Function

'---------------------------------------------------------------
' Devuelve la fecha menor de dos
'---------------------------------------------------------------
Private Function FechaMenor(ByVal fecha1 As Variant, ByVal fecha2 As Variant) As Variant
    If Not IsDate(fecha1) And Not IsDate(fecha2) Then
        FechaMenor = ""
    ElseIf Not IsDate(fecha1) Then
        FechaMenor = fecha2
    ElseIf Not IsDate(fecha2) Then
        FechaMenor = fecha1
    ElseIf CDate(fecha1) < CDate(fecha2) Then
        FechaMenor = fecha1
    Else
        FechaMenor = fecha2
    End If
End Function

'---------------------------------------------------------------
' Devuelve la fecha mayor de dos
'---------------------------------------------------------------
Private Function FechaMayor(ByVal fecha1 As Variant, ByVal fecha2 As Variant) As Variant
    If Not IsDate(fecha1) And Not IsDate(fecha2) Then
        FechaMayor = ""
    ElseIf Not IsDate(fecha1) Then
        FechaMayor = fecha2
    ElseIf Not IsDate(fecha2) Then
        FechaMayor = fecha1
    ElseIf CDate(fecha1) > CDate(fecha2) Then
        FechaMayor = fecha1
    Else
        FechaMayor = fecha2
    End If
End Function

'---------------------------------------------------------------
' Obtiene la suma de una columna SOLO para las �rdenes v�lidas
'---------------------------------------------------------------
Private Function ObtenerValorOrdenesValidas(ByVal hojaOrigen As Worksheet, _
                                             ByVal listaOrdenes As String, _
                                             ByVal columna As String) As Variant
    Dim arrOrdenes As Variant
    Dim i As Long
    Dim fila As Long
    Dim suma As Double
    Dim orden As String
    Dim ultimaFila As Long
    
    suma = 0
    arrOrdenes = Split(listaOrdenes, ",")
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, "A").End(xlUp).Row
    
    For i = LBound(arrOrdenes) To UBound(arrOrdenes)
        orden = Trim(arrOrdenes(i))
        
        ' Buscar la fila de esta orden
        For fila = 2 To ultimaFila
            If Trim(CStr(hojaOrigen.Cells(fila, "A").Value)) = orden Then
                If IsNumeric(hojaOrigen.Cells(fila, columna).Value) Then
                    suma = suma + CDbl(hojaOrigen.Cells(fila, columna).Value)
                End If
                Exit For
            End If
        Next fila
    Next i
    
    If suma = 0 Then
        ObtenerValorOrdenesValidas = ""
    Else
        ObtenerValorOrdenesValidas = suma
    End If
End Function

'---------------------------------------------------------------
' Obtiene la fecha M�NIMA de una columna SOLO para las �rdenes v�lidas
'---------------------------------------------------------------
Private Function ObtenerFechaMinOrdenesValidas(ByVal hojaOrigen As Worksheet, _
                                                ByVal listaOrdenes As String, _
                                                ByVal columna As String) As Variant
    Dim arrOrdenes As Variant
    Dim i As Long
    Dim fila As Long
    Dim fechaMin As Variant
    Dim fechaActual As Variant
    Dim orden As String
    Dim ultimaFila As Long
    
    fechaMin = Empty
    arrOrdenes = Split(listaOrdenes, ",")
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, "A").End(xlUp).Row
    
    For i = LBound(arrOrdenes) To UBound(arrOrdenes)
        orden = Trim(arrOrdenes(i))
        
        For fila = 2 To ultimaFila
            If Trim(CStr(hojaOrigen.Cells(fila, "A").Value)) = orden Then
                fechaActual = hojaOrigen.Cells(fila, columna).Value
                If IsDate(fechaActual) Then
                    If IsEmpty(fechaMin) Then
                        fechaMin = fechaActual
                    ElseIf CDate(fechaActual) < CDate(fechaMin) Then
                        fechaMin = fechaActual
                    End If
                End If
                Exit For
            End If
        Next fila
    Next i
    
    ObtenerFechaMinOrdenesValidas = fechaMin
End Function

'---------------------------------------------------------------
' Obtiene la fecha M�XIMA de una columna SOLO para las �rdenes v�lidas
'---------------------------------------------------------------
Private Function ObtenerFechaMaxOrdenesValidas(ByVal hojaOrigen As Worksheet, _
                                                ByVal listaOrdenes As String, _
                                                ByVal columna As String) As Variant
    Dim arrOrdenes As Variant
    Dim i As Long
    Dim fila As Long
    Dim fechaMax As Variant
    Dim fechaActual As Variant
    Dim orden As String
    Dim ultimaFila As Long
    
    fechaMax = Empty
    arrOrdenes = Split(listaOrdenes, ",")
    ultimaFila = hojaOrigen.Cells(hojaOrigen.Rows.Count, "A").End(xlUp).Row
    
    For i = LBound(arrOrdenes) To UBound(arrOrdenes)
        orden = Trim(arrOrdenes(i))
        
        For fila = 2 To ultimaFila
            If Trim(CStr(hojaOrigen.Cells(fila, "A").Value)) = orden Then
                fechaActual = hojaOrigen.Cells(fila, columna).Value
                If IsDate(fechaActual) Then
                    If IsEmpty(fechaMax) Then
                        fechaMax = fechaActual
                    ElseIf CDate(fechaActual) > CDate(fechaMax) Then
                        fechaMax = fechaActual
                    End If
                End If
                Exit For
            End If
        Next fila
    Next i
    
    ObtenerFechaMaxOrdenesValidas = fechaMax
End Function

'---------------------------------------------------------------
' Limpia registros del hist�rico con m�s de 3 meses de antig�edad
' Ahora busca en columna B (FECHA ENV�O)
'---------------------------------------------------------------
Public Sub LimpiarHistoricoAntiguo(ByVal nombreHojaHistorico As String)
    Dim wsHistorico As Worksheet
    Dim ultimaFila As Long
    Dim fila As Long
    Dim fechaEnvio As Date
    Dim fechaLimite As Date
    Dim registrosEliminados As Long
    Dim estabaProtegida As Boolean
    
    On Error Resume Next
    Set wsHistorico = ThisWorkbook.Worksheets(nombreHojaHistorico)
    On Error GoTo 0
    
    If wsHistorico Is Nothing Then Exit Sub
    
    ' Verificar si la hoja est� protegida y desprotegerla
    estabaProtegida = wsHistorico.ProtectContents
    If estabaProtegida Then
        On Error Resume Next
        wsHistorico.Unprotect password:=""
        If wsHistorico.ProtectContents Then
            wsHistorico.Unprotect password:=ModuloConfigSegura.ObtenerPasswordLOG()
        End If
        On Error GoTo 0
    End If
    
    ' Fecha l�mite:   3 meses atr�s desde hoy
    fechaLimite = DateAdd("m", -3, Date)
    
    ultimaFila = wsHistorico.Cells(wsHistorico.Rows.Count, "A").End(xlUp).Row
    registrosEliminados = 0
    
    ' Recorrer desde abajo hacia arriba para evitar problemas al eliminar
    For fila = ultimaFila To 2 Step -1
        If IsDate(wsHistorico.Cells(fila, "B").Value) Then
            fechaEnvio = CDate(wsHistorico.Cells(fila, "B").Value)
            
            ' Si la fecha de env�o es anterior al l�mite, eliminar fila
            If fechaEnvio < fechaLimite Then
                wsHistorico.Rows(fila).Delete
                registrosEliminados = registrosEliminados + 1
            End If
        End If
    Next fila
    
    ' Reproteger la hoja si estaba protegida
    If estabaProtegida Then
        On Error Resume Next
        wsHistorico.Protect password:=ModuloConfigSegura.ObtenerPasswordLOG(), UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    
End Sub

'---------------------------------------------------------------
' Verifica qu� N� ORDEN de una lista ya existen en el hist�rico
' Devuelve una cadena con los duplicados encontrados
'---------------------------------------------------------------
Public Function ObtenerOrdenesDuplicadas(ByVal listaOrdenes As String, _
                                          ByVal nombreHojaHistorico As String) As String
    Dim ordenes As Variant
    Dim i As Long
    Dim duplicados As String
    
    ordenes = Split(listaOrdenes, ",")
    duplicados = ""
    
    For i = LBound(ordenes) To UBound(ordenes)
        If OrdenYaEnviada(Trim(ordenes(i)), nombreHojaHistorico) Then
            If duplicados = "" Then
                duplicados = Trim(ordenes(i))
            Else
                duplicados = duplicados & ", " & Trim(ordenes(i))
            End If
        End If
    Next i
    
    ObtenerOrdenesDuplicadas = duplicados
End Function

'---------------------------------------------------------------
' Confirmar reenv�o si hay duplicados
'---------------------------------------------------------------
Public Function ConfirmarReenvio(ByVal ordenesDuplicadas As String) As Boolean
    Dim resp As VbMsgBoxResult

    resp = MsgBox( _
        "Los siguientes N� ORDEN ya fueron enviados a lavander�a anteriormente:" & vbCrLf & vbCrLf & _
        ordenesDuplicadas & vbCrLf & vbCrLf & _
        "�Desea continuar con el env�o?", _
        vbQuestion + vbYesNo, _
        "�rdenes ya enviadas")

    ConfirmarReenvio = (resp = vbYes)
End Function

Public Function EsAlojamientoValidoGijon(ByVal texto As String) As Boolean
    Dim validos As Variant
    validos = Array("1", "2", "3", "4", "5", "6", "7", _
                    "OF.1", "OF.2", "OF.3", _
                    "EST.1", "EST.2", "EST.3")

    EsAlojamientoValidoGijon = ListaAlojamientosValida(texto, validos)
End Function

Public Function EsAlojamientoValidoOviedo(ByVal texto As String) As Boolean
    Dim validos As Variant
    validos = Array("1", "2", "3", "4", "5", "6", "7", "8", _
                    "9", "10", "11", "12", "13", "14", "15")

    EsAlojamientoValidoOviedo = ListaAlojamientosValida(texto, validos)
End Function

'---------------------------------------------------------------
' FUNCIONES PARA GIJ�N (misma l�gica que Soto)
' Ya est�n cubiertas por las funciones gen�ricas que reciben
' el nombre de la hoja hist�rico como par�metro
'---------------------------------------------------------------
' No es necesario a�adir funciones adicionales, ya que:
' - OrdenYaEnviada(numOrden, "HISTORICO_LAVANDERIA_GIJON")
' - RegistrarEnvioEnHistorico(hoja, filaTotal, listaOrdenes, "HISTORICO_LAVANDERIA_GIJON")
' - LimpiarHistoricoAntiguo("HISTORICO_LAVANDERIA_GIJON")
' - ObtenerOrdenesDuplicadas(listaOrdenes, "HISTORICO_LAVANDERIA_GIJON")
' Todas funcionan pasando el nombre de la hoja hist�rico correspondiente
