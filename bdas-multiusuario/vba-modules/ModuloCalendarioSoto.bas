Attribute VB_Name = "ModuloCalendarioSoto"
'Attribute VB_Name = "ModuloCalendarioSoto"
Option Explicit

'=================================================================================
' Modulo: ModuloCalendarioSoto
' Version: 17.0 - Sistema de Cache por Fecha de Corte
' Fecha: 2026-01-31
'=================================================================================

' VARIABLE PRIVADA para almacenar informaci?n de solapes entre reservas
Private tipoSolape As Object

' Procedimiento para inicializar/reiniciar el diccionario de solapes
Private Sub DefinirTipoSolape()
    On Error Resume Next
    If tipoSolape Is Nothing Then
        Set tipoSolape = CreateObject("Scripting.Dictionary")
    Else
        tipoSolape.RemoveAll  ' Limpiar datos anteriores
    End If
End Sub

'=================================================================================
' INTERFACES PUBLICAS - Puntos de entrada principales
'=================================================================================

' Punto de entrada 1: Actualizaci?n COMPLETA del calendario (hist?rico + futuro)
Public Sub ActualizarCalendarioCompletoSoto()
    ' FLUJO: Llamado desde bot?n Ribbon - Actualiza TODO el calendario
    ' OBJETIVO: Refrescar completo incluyendo datos hist?ricos
    ActualizarCalendarioInterno False, False  ' Par?metros: NO solo futuras, NO usar cache
    GestionCacheCalendarios.MarcarCacheSotoCargado  ' Marcar cache como actualizado
    Application.StatusBar = "Calendario Soto: Historico completo cargado y en cache"
    Call ResaltarFestivos  ' Funcionalidad adicional externa
End Sub

' Punto de entrada 2: Actualizaci?n PARCIAL del calendario (solo futuras)
Public Sub ActualizarCalendarioFuturoSoto()
    ' FLUJO: Llamado al activar la hoja - Solo actualiza fechas futuras
    ' OBJETIVO: Actualizaci?n r�pida respetando cache de datos hist?ricos
    ActualizarCalendarioInterno True, GestionCacheCalendarios.CacheSotoCargado()
    ' Par?metros: SI solo futuras, SI usar cache si est� disponible
End Sub

'=================================================================================
' PROCEDIMIENTO PRINCIPAL - N?cleo de toda la l�gica
'=================================================================================

' Procedimiento PRIVADO principal que orquesta toda la actualizaci?n
Private Sub ActualizarCalendarioInterno(ByVal soloFuturas As Boolean, ByVal usarCache As Boolean)
    ' VARIABLES PRINCIPALES:
    Dim wsRes As Worksheet, wsCal As Worksheet, wsBloq As Worksheet  ' Referencias a hojas
    Dim dateMapping As Object  ' Diccionario: Fecha ? Columna en calendario
    Dim reservasProcessadas As Collection  ' Reservas procesadas en memoria
    Dim bloqueosProcessados As Collection  ' Bloqueos procesados en memoria
    Dim mapaOcupacion As Object  ' Diccionario para controlar ocupaci�n (evitar solapes)
    Dim huboSolape As Boolean  ' Bandera: ?se detectaron solapes?
    Dim reservasOcultas As Long  ' Contador de reservas en apartamentos ocultos
    Dim fechaCorte As Date  ' Fecha l?mite para actualizaci?n parcial (cache)
    Dim colCorte As Long  ' Columna correspondiente a fechaCorte en calendario
    Dim estabaProtegida As Boolean  ' Detectar si la hoja del calendario est� protegida
    
    On Error GoTo ErrorHandler  ' Manejo estructurado de errores
    
    ' OBTENER REFERENCIAS a las 3 hojas clave:
    ' - RESIDENCIA SOTO: Datos de reservas
    ' - Calendario SOTO: Calendario visual a actualizar
    ' - Apartamentos Bloqueados SOTO: Datos de bloqueos
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA SOTO")
    Set wsCal = ThisWorkbook.Worksheets("Calendario SOTO")
    Set wsBloq = ThisWorkbook.Worksheets("Apartamentos Bloqueados SOTO")
    On Error GoTo ErrorHandler
    
    ' VALIDACIN CR?TICA: Verificar que existen las 3 hojas necesarias
    If wsRes Is Nothing Or wsCal Is Nothing Or wsBloq Is Nothing Then
        MsgBox "ERROR: No se encontraron las hojas necesarias para SOTO.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' === NUEVO: Sincronizar datos desde Access DB antes de pintar el calendario ===
    modDatabase.SincronizarHojaDesdeAccess "SOTO", wsRes
    
    ' Desproteger la hoja del calendario temporalmente para realizar la limpieza y pintado sin errores
    estabaProtegida = wsCal.ProtectContents
    If estabaProtegida Then wsCal.Unprotect password:=""
    
    ' OPTIMIZACIN: Desactivar caracter?sticas de Excel para acelerar el proceso
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    
    ' PASO 1: Crear mapeo de fechas a columnas en el calendario
    ' IMPORTANTE: Cada fecha ocupa 3 columnas en el calendario (tramo 0,1,2)
    Set dateMapping = CrearMapeoFechas(wsCal)
    If dateMapping Is Nothing Or dateMapping.Count = 0 Then
        MsgBox "ERROR: No se pudo crear mapeo de fechas.", vbExclamation, "Error"
        GoTo CleanExit
    End If
    
    ' PASO 2: Determinar fecha y columna de corte para actualizaci?n parcial (cache)
    ' L?GICA: Si hay cache, solo actualizar desde una fecha espec?fica hacia adelante
    If usarCache Then
        fechaCorte = GestionCacheCalendarios.ObtenerFechaCorteSoto()  ' Obtener fecha del cache
        If fechaCorte = 0 Then fechaCorte = Date  ' Si no hay cache, usar fecha actual
        fechaCorte = fechaCorte - 3  ' Margen de 3 d�as para evitar desfases
    Else
        fechaCorte = 0  ' Sin corte = actualizar todo
    End If
    
    ' Buscar columna correspondiente a la fecha de corte en el calendario
    colCorte = 0
    If fechaCorte > 0 Then
        If dateMapping.Exists(CLng(fechaCorte)) Then
            colCorte = dateMapping(CLng(fechaCorte))
        Else
            ' Si no encuentra exacta, buscar la fecha m�s cercana (hasta 10 d�as despu�s)
            Dim fechaBuscar As Date
            For fechaBuscar = fechaCorte To fechaCorte + 10
                If dateMapping.Exists(CLng(fechaBuscar)) Then
                    colCorte = dateMapping(CLng(fechaBuscar))
                    Exit For
                End If
            Next fechaBuscar
        End If
    End If
    
    ' PASO 3: LIMPIAR el calendario seg�n el modo (completo o parcial)
    ' DECISI?N CR?TICA: Seg?n cache y columna de corte, elegir estrategia de limpieza
    If usarCache And colCorte > 0 Then
        ' MODO CACHE: Limpiar solo desde colCorte hacia adelante
        LimpiarCalendarioParcial wsCal, colCorte  ' Limpieza general
        LimpiarBloqueosParcial wsCal, colCorte    ' Limpieza ESPEC?FICA de bloqueos (NUEVO - soluci?n al problema)
    Else
        ' MODO COMPLETO: Limpiar todo el calendario
        LimpiarCalendarioRapido wsCal  ' Limpieza general completa
        LimpiarBloqueos wsCal          ' Limpieza espec?fica de bloqueos completa
    End If
    
    ' PASO 4: PROCESAR DATOS en memoria para mayor eficiencia
    ' PROCESAR RESERVAS: Leer datos de RESIDENCIA SOTO y convertirlos a estructura en memoria
    Set reservasProcessadas = ProcesarReservasEnMemoriaV13(wsRes, dateMapping, soloFuturas, fechaCorte)
    
    ' PROCESAR BLOQUEOS: Leer datos de Apartamentos Bloqueados SOTO y convertirlos a estructura en memoria
    Set bloqueosProcessados = ProcesarBloqueosEnMemoriaV13(wsBloq, dateMapping, soloFuturas, fechaCorte)
    
    ' PASO 5: Contar reservas en apartamentos OCULTOS (para mostrar advertencia)
    Dim listaOcultas As String: listaOcultas = ""
    reservasOcultas = ContarReservasEnApartamentosOcultos(wsRes, wsCal, listaOcultas)
    
    ' PASO 6: Inicializar estructuras para DETECCI?N DE SOLAPES
    DefinirTipoSolape  ' Inicializar diccionario de solapes
    Set mapaOcupacion = CreateObject("Scripting.Dictionary")  ' Diccionario para control ocupaci�n
    
    ' PASO 7: APLICAR reservas y bloqueos al calendario (PINTADO PRINCIPAL)
    ' Esta funci?n es la m�s importante: pinta colores y textos en el calendario
    huboSolape = AplicarReservasYBloqueosConTramosV16(wsCal, reservasProcessadas, bloqueosProcessados, mapaOcupacion, tipoSolape)
    
    ' PASO 8: REPINTAR reservas PAGADAS con color verde especial
    RepintarReservasPagadas wsRes, wsCal, reservasProcessadas
    
    ' PASO 9: Aplicar formato BASE al calendario (alineaci?n, altos, anchos)
    AplicarFormatoBaseCalendario wsCal
    
    ' PASO 10: Mostrar ADVERTENCIAS si hay reservas ocultas o solapes
        ' Advertencia 1: Reservas en apartamentos ocultos
    If reservasOcultas > 0 Then
        Dim mensajeOcultas As String
        If reservasOcultas = 1 Then
            mensajeOcultas = "ATENCION: Hay 1 reserva actual/futura asignada a un apartamento oculto (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Esta reserva NO se muestra en el calendario." & vbCrLf & vbCrLf & _
                           "Para verla, debes mostrar la fila correspondiente en 'Calendario SOTO'."
        Else
            mensajeOcultas = "ATENCION: Hay " & reservasOcultas & " reservas actuales/futuras asignadas a apartamentos ocultos (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Estas reservas NO se muestran en el calendario." & vbCrLf & vbCrLf & _
                           "Para verlas, debes mostrar las filas correspondientes en 'Calendario SOTO'."
        End If
        MsgBox mensajeOcultas, vbExclamation, "Reservas en Apartamentos Ocultos"
    End If
    
    ' Advertencia 2: Solapes detectados (reservas que se superponen)
    If huboSolape Then
        MostrarDetalleSolapes  ' Mostrar detalles espec�ficos de los solapes
    End If

CleanExit:
    ' RESTAURAR configuraci�n de Excel (optimizaci�n)
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    Exit Sub

ErrorHandler:
    ' RESTAURAR configuraci�n incluso en caso de error
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "ERROR: " & Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' LIMPIEZA DE CALENDARIO - Versiones optimizadas
'=================================================================================

' Limpieza COMPLETA del calendario (sin cache)
Private Sub LimpiarCalendarioRapido(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long
    Dim rango As Range
    
    ' Determinar dimensiones del �rea de datos del calendario
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub  ' No hay datos
    
    Set rango = wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol))
    
    ' PASO CR?TICO: Deshacer combinaciones de celdas
    ' IMPORTANTE: Las celdas combinadas pueden causar problemas en la limpieza
    On Error Resume Next
    rango.UnMerge
    On Error GoTo 0
    
    ' LIMPIAR contenido y formato:
    rango.ClearContents       ' Borrar textos (nombres, n� orden, "BLOQUEADA")
    rango.Interior.ColorIndex = xlNone  ' Quitar colores
    rango.Font.Size = 11      ' Restaurar tama?o de fuente por defecto
    rango.Font.Bold = False   ' Quitar negrita
    
    Set rango = Nothing  ' Liberar referencia
End Sub

' Limpieza PARCIAL del calendario (modo cache)
Private Sub LimpiarCalendarioParcial(wsCal As Worksheet, ByVal colDesde As Long)
    ' OBJETIVO: Limpiar solo desde colDesde hacia adelante (optimizaci?n con cache)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colDesde < 2 Then colDesde = 2  ' Validaci�n m?nima
    
    Dim ultimaFila As Long, ultimaCol As Long
    Dim rango As Range
    
    ' Determinar dimensiones
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    ' Validar que hay datos en el rango a limpiar
    If ultimaFila < 2 Or ultimaCol < colDesde Then Exit Sub
    
    ' Definir rango espec�fico (desde colDesde hasta �ltima columna)
    Set rango = wsCal.Range(wsCal.Cells(2, colDesde), wsCal.Cells(ultimaFila, ultimaCol))
    
    ' PASO CR?TICO: Deshacer combinaciones de celdas
    On Error Resume Next
    rango.UnMerge
    On Error GoTo 0
    
    ' LIMPIAR contenido y formato (solo en el rango parcial)
    rango.ClearContents
    rango.Interior.ColorIndex = xlNone
    rango.Font.Size = 11
    rango.Font.Bold = False
    
    Set rango = Nothing
End Sub

' Limpieza ESPEC�FICA de BLOQUEOS
Private Sub LimpiarBloqueos(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long, i As Long, j As Long
    Dim celda As Range
    Dim arrValues As Variant
    
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub
    
    ' Leer toda la cuadr�cula en memoria
    arrValues = wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol)).Value
    If Not IsArray(arrValues) Then Exit Sub
    
    For i = 2 To ultimaFila
        For j = 2 To ultimaCol
            Dim val As Variant
            val = arrValues(i - 1, j - 1)
            
            ' Solo consultamos el objeto Celda si contiene "BLOQUEADA" o si no est� vac�o
            If val = "BLOQUEADA" Or Not IsEmpty(val) Then
                Set celda = wsCal.Cells(i, j)
                If celda.Value = "BLOQUEADA" Or celda.Interior.color = RGB(255, 192, 0) Then
                    If celda.MergeCells Then
                        celda.MergeCells = False
                    End If
                    celda.Value = ""
                    celda.Interior.ColorIndex = xlNone
                    celda.Font.Size = 11
                    celda.Font.Bold = False
                End If
                Set celda = Nothing
            End If
        Next j
    Next i
End Sub

Private Sub LimpiarBloqueosParcial(wsCal As Worksheet, ByVal colDesde As Long)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colDesde < 2 Then colDesde = 2
    
    Dim ultimaFila As Long, ultimaCol As Long, i As Long, j As Long
    Dim celda As Range
    Dim arrValues As Variant
    
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaFila < 2 Or ultimaCol < colDesde Then Exit Sub
    
    ' Leer toda la cuadr�cula parcial en memoria
    arrValues = wsCal.Range(wsCal.Cells(2, colDesde), wsCal.Cells(ultimaFila, ultimaCol)).Value
    If Not IsArray(arrValues) Then Exit Sub
    
    For i = 2 To ultimaFila
        For j = colDesde To ultimaCol
            Dim val As Variant
            val = arrValues(i - 1, j - colDesde + 1)
            
            ' Solo consultamos el objeto Celda si contiene "BLOQUEADA" o si no est� vac�o
            If val = "BLOQUEADA" Or Not IsEmpty(val) Then
                Set celda = wsCal.Cells(i, j)
                If celda.Value = "BLOQUEADA" Or celda.Interior.color = RGB(255, 192, 0) Then
                    If celda.MergeCells Then
                        celda.MergeCells = False
                    End If
                    celda.Value = ""
                    celda.Interior.ColorIndex = xlNone
                    celda.Font.Size = 11
                    celda.Font.Bold = False
                End If
                Set celda = Nothing
            End If
        Next j
    Next i
End Sub

'=================================================================================
' FORMATO BASE DEL CALENDARIO - Configuraci?n visual
'=================================================================================

Private Sub AplicarFormatoBaseCalendario(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub
    
    ' APLICAR formato b?sico a todo el �rea de datos
    With wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol))
        ' NOTA: No se aplica Font.Bold ni Font.Size aqu? para no sobreescribir
        '       el formato espec�fico del N� ORDEN que se aplica despu�s
        .HorizontalAlignment = xlLeft    ' Alinear texto a la izquierda
        .VerticalAlignment = xlBottom    ' Alinear texto al fondo
        .WrapText = True                 ' Ajustar texto en varias l�neas
    End With
    
    ' AJUSTAR ALTURA de filas (solo filas pares = apartamentos)
    Dim fila As Long
    For fila = 2 To ultimaFila Step 2
        If Not wsCal.Rows(fila).Hidden Then
            wsCal.Rows(fila).RowHeight = 80  ' Alto fijo para celdas de apartamentos
        End If
    Next fila
    
    ' AJUSTAR ANCHO de columnas (tramo 2 de cada d?a)
    Dim col As Long
    For col = 2 To ultimaCol Step 3
        If col + 2 <= wsCal.Columns.Count Then
            wsCal.Columns(col + 2).ColumnWidth = 37  ' Ancho espec�fico para tramo 2
        End If
    Next col
End Sub

'=================================================================================
' CREAR MAPEO DE FECHAS - Conversi?n fecha ? columna
'=================================================================================

Private Function CrearMapeoFechas(wsCal As Worksheet) As Object
    ' OBJETIVO: Crear diccionario que mapea cada fecha (en formato Long) a su
    '           columna de inicio en el calendario (cada fecha ocupa 3 columnas)
    
    On Error Resume Next
    Set CrearMapeoFechas = Nothing
    If wsCal Is Nothing Then Exit Function
    
    Dim dateMapping As Object
    Set dateMapping = CreateObject("Scripting.Dictionary")
    
    Dim ultimaCol As Long, col As Long
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaCol < 2 Then Set CrearMapeoFechas = dateMapping: Exit Function
    
    ' RECORRER columnas en saltos de 3 (cada fecha ocupa 3 columnas)
    For col = 2 To ultimaCol Step 3
        If IsDate(wsCal.Cells(1, col).Value) Then
            Dim fechaLng As Long
            fechaLng = CLng(wsCal.Cells(1, col).Value)  ' Convertir fecha a Long
            If Not dateMapping.Exists(fechaLng) Then dateMapping.Add fechaLng, col
        End If
    Next col
    
    Set CrearMapeoFechas = dateMapping
End Function

'=================================================================================
' PROCESADO EN MEMORIA - Lectura eficiente de datos
'=================================================================================

' Procesar RESERVAS en memoria (optimizado con arrays)
Private Function ProcesarReservasEnMemoriaV13(wsRes As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    ' OBJETIVO: Leer TODAS las reservas de la hoja RESIDENCIA SOTO
    '           y convertirlas a una colecci?n de arrays en memoria
    
    Dim reservasProcessadas As Collection
    Set reservasProcessadas = New Collection
    On Error Resume Next
    
    ' Validaciones b?sicas
    If wsRes Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    End If
    
    ' Determinar �ltima fila con datos
    Dim lastRow As Long
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    ' LEER DATOS EN ARRAY (optimizaci?n cr?tica)
    ' Leer todo el rango A2:S[lastRow] de una vez en memoria
    Dim arrRes As Variant
    arrRes = wsRes.Range("A2:S" & lastRow).Value
    If Not IsArray(arrRes) Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    ' Variables locales
    Dim i As Long, room As String, arrRooms As Variant, r As Variant, reservaArray As Variant
    Dim hoy As Date: hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro seg�n par?metros
    If fechaCorte > 0 Then
        fechaFiltro = fechaCorte  ' Modo cache: filtrar por fecha de corte
    ElseIf soloFuturas Then
        fechaFiltro = hoy - 3     ' Solo futuras: filtrar con margen de 3 d�as
    Else
        fechaFiltro = 0           ' Sin filtro: procesar todo
    End If
    
    ' PROCESAR cada fila del array en memoria
    For i = 1 To UBound(arrRes, 1)
        ' CRITERIOS para considerar una reserva v�lida:
        ' 1. Tiene fecha de entrada (col 12/L) y salida (col 13/M)
        ' 2. Tiene apartamento asignado (col 19/S)
        If IsDate(arrRes(i, 12)) And IsDate(arrRes(i, 13)) And Not IsEmpty(arrRes(i, 19)) Then
            
            ' APLICAR FILTRO de fecha si corresponde
            If fechaFiltro > 0 Then
                If CDate(arrRes(i, 13)) < fechaFiltro Then GoTo NextI  ' Saltar reservas antiguas
            End If
            
            ' PROCESAR apartamento(s): puede ser uno solo o m�ltiples separados por coma
            room = Trim(CStr(arrRes(i, 19)))
            If InStr(room, ",") > 0 Then
                ' RESERVA MULTIPLE: Un N� Orden para varios apartamentos
                arrRooms = Split(room, ",")
                For Each r In arrRooms
                    If Trim(r) <> "" Then
                        ' Crear array de reserva para cada apartamento
                        reservaArray = CrearReservaArrayV12(arrRes, i, Trim(r), dateMapping)
                        If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
                    End If
                Next r
            Else
                ' RESERVA SIMPLE: Un N� Orden para un apartamento
                reservaArray = CrearReservaArrayV12(arrRes, i, room, dateMapping)
                If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
            End If
        End If
NextI:
    Next i
    
    Set ProcesarReservasEnMemoriaV13 = reservasProcessadas
End Function

' Procesar BLOQUEOS en memoria (similar a reservas)
Private Function ProcesarBloqueosEnMemoriaV13(wsBloq As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    ' OBJETIVO: Leer TODOS los bloqueos de la hoja Apartamentos Bloqueados SOTO
    '           y convertirlos a una colecci?n de arrays en memoria
    
    Dim bloqueosProcessados As Collection
    Set bloqueosProcessados = New Collection
    On Error Resume Next
    
    ' Validaciones b?sicas
    If wsBloq Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    End If
    
    ' Determinar �ltima fila con datos
    Dim lastRow As Long
    lastRow = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    ' LEER DATOS EN ARRAY
    Dim arrBloq As Variant
    arrBloq = wsBloq.Range("A2:C" & lastRow).Value
    If Not IsArray(arrBloq) Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    ' Variables locales
    Dim i As Long, bloqueoArray As Variant, hoy As Date
    hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro (misma l�gica que reservas)
    If fechaCorte > 0 Then
        fechaFiltro = fechaCorte
    ElseIf soloFuturas Then
        fechaFiltro = hoy - 3
    Else
        fechaFiltro = 0
    End If
    
    ' PROCESAR cada fila del array en memoria
    For i = 1 To UBound(arrBloq, 1)
        ' CRITERIOS para considerar un bloqueo v�lido:
        ' 1. Tiene nombre de apartamento (col A)
        ' 2. Tiene fecha inicio (col B) y fin (col C)
        If arrBloq(i, 1) <> "" And IsDate(arrBloq(i, 2)) And IsDate(arrBloq(i, 3)) Then
            
            ' APLICAR FILTRO de fecha si corresponde
            If fechaFiltro > 0 Then
                If CDate(arrBloq(i, 3)) < fechaFiltro Then GoTo NextJ
            End If
            
            ' Crear array de bloqueo
            bloqueoArray = CrearBloqueoArrayV12(arrBloq, i, dateMapping)
            If bloqueoArray(0) > 0 Then bloqueosProcessados.Add bloqueoArray
        End If
NextJ:
    Next i
    
    Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados
End Function

'=================================================================================
' CREAR ARRAYS DE RESERVA Y BLOQUEO - Estructuras en memoria
'=================================================================================

' Crear array de RESERVA (estructura: [fila, colInicio, colFin, texto])
Private Function CrearReservaArrayV12(arrRes As Variant, i As Long, ByVal room As String, dateMapping As Object) As Variant
    ' ESTRUCTURA del array de reserva:
    ' [0] = Fila en calendario (2,4,6,...,18)
    ' [1] = Columna de inicio (check-in)
    ' [2] = Columna de fin (check-out)
    ' [3] = Texto a mostrar (nombre + N� Orden)
    
    Dim reservaArray(0 To 3) As Variant
    reservaArray(0) = 0: reservaArray(1) = 0: reservaArray(2) = 0: reservaArray(3) = ""
    On Error Resume Next
    
    ' CONVERTIR nombre de apartamento a n�mero de fila en calendario
    Select Case room
        Case "Ap.2": reservaArray(0) = 2
        Case "Ap.4": reservaArray(0) = 4
        Case "Ap.5": reservaArray(0) = 6
        Case "Ap.6": reservaArray(0) = 8
        Case "Ap.7": reservaArray(0) = 10
        Case "Ap.8": reservaArray(0) = 12
        Case "Ap.9": reservaArray(0) = 14
        Case "Ap.10": reservaArray(0) = 16
        Case "Ap.12": reservaArray(0) = 18
        Case Else: CrearReservaArrayV12 = reservaArray: Exit Function
    End Select
    
    ' VERIFICAR si el apartamento est� OCULTO en el calendario
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario SOTO")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(reservaArray(0)).Hidden Then
            reservaArray(0) = 0  ' Marcar como inv?lida
            CrearReservaArrayV12 = reservaArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ' VALIDAR fechas
    If Not IsDate(arrRes(i, 12)) Or Not IsDate(arrRes(i, 13)) Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    Dim checkin As Date, checkout As Date
    checkin = CDate(arrRes(i, 12))
    checkout = CDate(arrRes(i, 13))
    
    ' CONVERTIR fechas a columnas usando el dateMapping
    If dateMapping.Exists(CLng(checkin)) Then reservaArray(1) = dateMapping(CLng(checkin))
    If dateMapping.Exists(CLng(checkout)) Then reservaArray(2) = dateMapping(CLng(checkout))
    
    ' Validar que se encontraron ambas columnas
    If reservaArray(1) = 0 Or reservaArray(2) = 0 Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    ' CONSTRUIR TEXTO a mostrar en calendario
    Dim fullName As String, displayName As String, pax As String
    fullName = Trim(CStr(arrRes(i, 11)))  ' Nombre (col K)
    pax = Trim(CStr(arrRes(i, 15)))       ' N� personas (col O)
    
    If fullName = "" Then
        ' Si no hay nombre, mostrar solo N� Orden
        reservaArray(3) = arrRes(i, 1)  ' N� Orden (col A)
    Else
        ' Formatear nombre: truncar si es muy largo
        If Len(fullName) <= 15 Then
            displayName = fullName & " (" & pax & ")"
        Else
            displayName = Left(fullName, 15) & ". (" & pax & ")"
        End If
        ' Texto final: nombre + salto de l�nea + N� Orden
        reservaArray(3) = displayName & vbNewLine & arrRes(i, 1)
    End If
    
    ' A?adir "C" si es comisi?n o destino
    If arrRes(i, 4) = "Comisi�n NO indem." Or arrRes(i, 4) = "Destino" Or arrRes(i, 4) = "Comisi�n" Then
        reservaArray(3) = reservaArray(3) & " C"
    End If
    
    CrearReservaArrayV12 = reservaArray
End Function

' Crear array de BLOQUEO (estructura similar a reserva)
Private Function CrearBloqueoArrayV12(arrBloq As Variant, i As Long, dateMapping As Object) As Variant
    ' ESTRUCTURA del array de bloqueo:
    ' [0] = Fila en calendario (2,4,6,...,18)
    ' [1] = Columna de inicio
    ' [2] = Columna de fin
    ' [3] = Nombre del apartamento (ej: "Ap.2")
    
    Dim bloqueoArray(0 To 3) As Variant
    bloqueoArray(0) = 0: bloqueoArray(1) = 0: bloqueoArray(2) = 0
    bloqueoArray(3) = CStr(arrBloq(i, 1))  ' Nombre del apartamento
    On Error Resume Next
    
    ' CONVERTIR nombre de apartamento a n�mero de fila
    Select Case bloqueoArray(3)
        Case "Ap.2": bloqueoArray(0) = 2
        Case "Ap.4": bloqueoArray(0) = 4
        Case "Ap.5": bloqueoArray(0) = 6
        Case "Ap.6": bloqueoArray(0) = 8
        Case "Ap.7": bloqueoArray(0) = 10
        Case "Ap.8": bloqueoArray(0) = 12
        Case "Ap.9": bloqueoArray(0) = 14
        Case "Ap.10": bloqueoArray(0) = 16
        Case "Ap.12": bloqueoArray(0) = 18
        Case Else: CrearBloqueoArrayV12 = bloqueoArray: Exit Function
    End Select
    
    ' VERIFICAR si el apartamento est� OCULTO
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario SOTO")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(bloqueoArray(0)).Hidden Then
            bloqueoArray(0) = 0
            CrearBloqueoArrayV12 = bloqueoArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ' VALIDAR fechas
    If Not IsDate(arrBloq(i, 2)) Or Not IsDate(arrBloq(i, 3)) Then
        bloqueoArray(0) = 0: CrearBloqueoArrayV12 = bloqueoArray: Exit Function
    End If
    
    Dim dIni As Date, dFin As Date
    dIni = CDate(arrBloq(i, 2))
    dFin = CDate(arrBloq(i, 3))
    
    ' CONVERTIR fechas a columnas
    If dateMapping.Exists(CLng(dIni)) Then bloqueoArray(1) = dateMapping(CLng(dIni))
    If dateMapping.Exists(CLng(dFin)) Then bloqueoArray(2) = dateMapping(CLng(dFin))
    
    ' Validar que se encontraron ambas columnas
    If bloqueoArray(1) = 0 Or bloqueoArray(2) = 0 Then bloqueoArray(0) = 0
    
    CrearBloqueoArrayV12 = bloqueoArray
End Function

'=================================================================================
' APLICAR RESERVAS Y BLOQUEOS - PINTADO en calendario
'=================================================================================

Private Function AplicarReservasYBloqueosConTramosV16(wsCal As Worksheet, reservas As Collection, bloqueos As Collection, mapaOcupacion As Object, solapeRecolector As Object) As Boolean
    ' OBJETIVO PRINCIPAL: Pintar colores y textos en el calendario
    ' L?GICA DE TRAMOS: Cada d?a ocupa 3 columnas:
    '   - Tramo 0: Check-out (columna base)
    '   - Tramo 1: D?as intermedios (columna base + 1)
    '   - Tramo 2: Check-in (columna base + 2) - Aqu? va el texto
    
    Dim huboSolape As Boolean: huboSolape = False  ' Bandera de solapes
    Dim i As Long, reservaArray As Variant, col As Long, clave As String
    Dim tramo As Long, bloqueoArray As Variant
    Dim colCheckin As Long, colCheckout As Long
    
    On Error Resume Next

    ' --- PINTAR RESERVAS ---
    For i = 1 To reservas.Count
        reservaArray = reservas(i)
        
        colCheckin = reservaArray(1)   ' Columna de inicio (check-in)
        colCheckout = reservaArray(2)  ' Columna de fin (check-out)
        
        ' DIA DE CHECK-IN - Solo TRAMO 3 (columna +2)
        ' Esta celda lleva el texto completo de la reserva
        clave = reservaArray(0) & "_" & (colCheckin + 2)
        With wsCal.Cells(reservaArray(0), colCheckin + 2)
            If .MergeCells Then .MergeCells = False  ' Evitar celdas combinadas
            
            ' VERIFICAR SOLAPE: Si ya est� ocupada esta celda
            If mapaOcupacion.Exists(clave) Then
                .Interior.color = vbRed: huboSolape = True  ' SOLAPE = ROJO
                RegistrarSolape wsCal, colCheckin + 2, reservaArray, mapaOcupacion(clave), solapeRecolector
            Else
                .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)  ' RESERVA = AMARILLO
            End If
            
            ' ESCRIBIR TEXTO de la reserva
            .Value = reservaArray(3)
            .Font.Size = 20      ' Tama?o grande para nombre
            .Font.Bold = False   ' Sin negrita para nombre
            
            ' FORMATEAR N� ORDEN: Tama?o m�s grande y negrita
            If InStr(.Value, vbNewLine) > 0 Then
                Dim lineBreakPos As Long
                lineBreakPos = InStr(.Value, vbNewLine)  ' Posici?n del salto de l�nea
                
                ' Aplicar formato solo al N� Orden (despu�s del salto)
                With .Characters(Start:=lineBreakPos + 1, Length:=Len(.Value) - lineBreakPos).Font
                    .Size = 32    ' Tama?o muy grande para N� Orden
                    .Bold = True  ' Negrita para N� Orden
                End With
            End If
        End With
        
        ' DIAS INTERMEDIOS - TODOS los tramos (0, 1, 2)
        ' Estas celdas solo llevan color, no texto
        If colCheckout > colCheckin + 3 Then
            For col = colCheckin + 3 To colCheckout - 3 Step 3
                For tramo = 0 To 2
                    clave = reservaArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(reservaArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        
                        ' Verificar solape para cada tramo
                        If mapaOcupacion.Exists(clave) Then
                            .Interior.color = vbRed: huboSolape = True
                            RegistrarSolape wsCal, col + tramo, reservaArray, mapaOcupacion(clave), solapeRecolector
                        Else
                            .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)
                        End If
                    End With
                Next tramo
            Next col
        End If
        
        ' DIA DE CHECK-OUT - Solo TRAMO 1 (columna +0)
        ' Solo color, no texto
        If colCheckout > colCheckin Then
            clave = reservaArray(0) & "_" & colCheckout
            With wsCal.Cells(reservaArray(0), colCheckout)
                If .MergeCells Then .MergeCells = False
                If mapaOcupacion.Exists(clave) Then
                    .Interior.color = vbRed: huboSolape = True
                    RegistrarSolape wsCal, colCheckout, reservaArray, mapaOcupacion(clave), solapeRecolector
                Else
                    .Interior.color = vbYellow: mapaOcupacion(clave) = reservaArray(3)
                End If
            End With
        End If
        
    Next i

    ' --- PINTAR BLOQUEOS ---
    ' Bloqueos pintados igual que reservas (intervalo semiabierto [inicio, fin))
    ' - Dia inicio: solo tramo checkin (col+2)
    ' - Dias intermedios: los 3 tramos
    ' - Dia fin: solo tramo checkout (col+0), si es distinto al dia inicio
    For i = 1 To bloqueos.Count
        bloqueoArray = bloqueos(i)
        
        ' DIA DE INICIO DEL BLOQUEO - solo tramo checkin (col+2)
        clave = bloqueoArray(0) & "_" & (bloqueoArray(1) + 2)
        With wsCal.Cells(bloqueoArray(0), bloqueoArray(1) + 2)
            If .MergeCells Then .MergeCells = False
            If mapaOcupacion.Exists(clave) Then
                .Interior.color = vbRed: huboSolape = True
                Dim numOrdenSolape As String
                numOrdenSolape = "-"
                If mapaOcupacion(clave) <> "BLOQUEO" Then
                    numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                End If
                RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(1) + 2, numOrdenSolape, bloqueoArray, solapeRecolector
            Else
                .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
            End If
            .Value = "BLOQUEADA"
        End With
        
        ' DIAS INTERMEDIOS DEL BLOQUEO - los 3 tramos
        If bloqueoArray(2) > bloqueoArray(1) + 3 Then
            For col = bloqueoArray(1) + 3 To bloqueoArray(2) - 3 Step 3
                For tramo = 0 To 2
                    clave = bloqueoArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(bloqueoArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        If mapaOcupacion.Exists(clave) Then
                            .Interior.color = vbRed: huboSolape = True
                            numOrdenSolape = "-"
                            If mapaOcupacion(clave) <> "BLOQUEO" Then
                                numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                            End If
                            RegistrarSolapeReservaBloqueo wsCal, col + tramo, numOrdenSolape, bloqueoArray, solapeRecolector
                        Else
                            .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
                        End If
                    End With
                Next tramo
                wsCal.Cells(bloqueoArray(0), col + 2).Value = "BLOQUEADA"
            Next col
        End If
        
        ' DIA DE FIN DEL BLOQUEO - solo tramo checkout (col+0), si es distinto al inicio
        If bloqueoArray(2) > bloqueoArray(1) Then
            clave = bloqueoArray(0) & "_" & bloqueoArray(2)
            With wsCal.Cells(bloqueoArray(0), bloqueoArray(2))
                If .MergeCells Then .MergeCells = False
                If mapaOcupacion.Exists(clave) Then
                    .Interior.color = vbRed: huboSolape = True
                    numOrdenSolape = "-"
                    If mapaOcupacion(clave) <> "BLOQUEO" Then
                        numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                    End If
                    RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(2), numOrdenSolape, bloqueoArray, solapeRecolector
                Else
                    .Interior.color = RGB(255, 192, 0): mapaOcupacion(clave) = "BLOQUEO"
                End If
            End With
        End If
    Next i
    
    AplicarReservasYBloqueosConTramosV16 = huboSolape
End Function

'=================================================================================
' REPINTAR RESERVAS PAGADAS - Color verde especial
'=================================================================================

Private Sub RepintarReservasPagadas(wsRes As Worksheet, wsCal As Worksheet, reservasProcessadas As Collection)
    ' OBJETIVO: Cambiar a color VERDE las reservas que est�n marcadas como "PAGADAS"
    '           en la columna AA (27) de RESIDENCIA SOTO
    
    On Error Resume Next
    
    If reservasProcessadas Is Nothing Then Exit Sub
    If reservasProcessadas.Count = 0 Then Exit Sub
    
    Dim i As Long, pagado As String, filaRes As Long, reservaArray As Variant
    Dim col As Long, tramo As Long, colCheckin As Long, colCheckout As Long
    
    For i = 1 To reservasProcessadas.Count
        reservaArray = reservasProcessadas(i)
        Dim partes() As String
        
        ' Extraer N� Orden del texto de la reserva
        If InStr(reservaArray(3), vbNewLine) > 0 Then
            partes = Split(reservaArray(3), vbNewLine)
            If UBound(partes) > 0 Then
                Dim numOrden As Variant
                numOrden = Trim(Split(partes(1), " ")(0))  ' Obtener N� Orden
                
                ' BUSCAR en RESIDENCIA SOTO si est� pagada
                filaRes = BuscarEnColumna(wsRes, 1, numOrden)  ' Buscar por N� Orden
                If filaRes > 0 Then
                    pagado = UCase(wsRes.Cells(filaRes, 27).Value)  ' Columna AA = Pagado
                    
                    ' Si est� pagado, repintar de VERDE
                    If MarcarSiPagadosEnResidencia.EsEstadoPagado(pagado) Then
                        colCheckin = reservaArray(1)
                        colCheckout = reservaArray(2)
                        
                        ' Check-in (tramo 2)
                        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = RGB(198, 239, 206)  ' VERDE CLARO
                        
                        ' D?as intermedios (todos los tramos)
                        If colCheckout > colCheckin + 3 Then
                            For col = colCheckin + 3 To colCheckout - 3 Step 3
                                For tramo = 0 To 2
                                    wsCal.Cells(reservaArray(0), col + tramo).Interior.color = RGB(198, 239, 206)
                                Next tramo
                            Next col
                        End If
                        
                        ' Check-out (tramo 0)
                        If colCheckout > colCheckin Then
                            wsCal.Cells(reservaArray(0), colCheckout).Interior.color = RGB(198, 239, 206)
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

'=================================================================================
' BUSQUEDAS Y UTILIDADES - Funciones auxiliares
'=================================================================================

' Buscar valor en columna espec?fica (optimizada)
Private Function BuscarEnColumna(ws As Worksheet, numCol As Long, valor As Variant) As Long
    On Error Resume Next
    BuscarEnColumna = 0
    If ws Is Nothing Then Exit Function
    If numCol < 1 Or numCol > ws.Columns.Count Then Exit Function
    
    Dim ultimaFila As Long, i As Long
    ultimaFila = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row
    If ultimaFila < 2 Then Exit Function
    
    ' B?squeda lineal (suficiente para tama?o de datos)
    For i = 2 To ultimaFila
        If Trim(CStr(ws.Cells(i, numCol).Value)) = Trim(CStr(valor)) Then
            BuscarEnColumna = i
            Exit Function
        End If
    Next i
End Function

'=================================================================================
' REGISTRO DE SOLAPES - Detecci?n y registro
'=================================================================================

Private Sub RegistrarSolape(wsCal As Worksheet, ByVal colIndex As Long, arr1 As Variant, arr2 As Variant, solapeRecolector As Object, Optional esBloqueo As Boolean = False)
    ' OBJETIVO: Registrar informaci?n detallada de solapes para mostrar despu�s
    
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, numOrdenReserva As String, numOrdenReserva2 As String
    Dim claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    ' Calcular columna del primer tramo del d?a
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    ' Obtener fecha de la cabecera del calendario
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    ' Obtener informaci?n de las reservas que solapan
    habStr = ObtenerNumApartamento(arr1(0))
    numOrdenReserva = ObtenerNumOrden(arr1(3))
    numOrdenReserva2 = ObtenerNumOrden(arr2)
    
    If numOrdenReserva = "-" Or numOrdenReserva2 = "-" Then Exit Sub
    
    ' Ordenar N?s Orden para clave ?nica
    Dim ord1 As Long, ord2 As Long
    If Not IsNumeric(numOrdenReserva) Or Not IsNumeric(numOrdenReserva2) Then Exit Sub
    
    ord1 = CLng(numOrdenReserva)
    ord2 = CLng(numOrdenReserva2)
    
    If ord1 > ord2 Then
        Dim temp As Long
        temp = ord1: ord1 = ord2: ord2 = temp
    End If
    
    ' Crear clave ?nica para el solape
    claveSolape = "AP_" & habStr & "_ORD_" & CStr(ord1) & "_" & CStr(ord2)
    
    ' Agregar o actualizar informaci?n en el diccionario
    If Not solapeRecolector.Exists(claveSolape) Then
        solapeRecolector.Add claveSolape, habStr & "|" & CStr(ord1) & "|" & CStr(ord2) & "|" & CStr(fechaLng) & ","
    Else
        solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
    End If
End Sub

' Registro espec�fico para solape RESERVA-BLOQUEO
Private Sub RegistrarSolapeReservaBloqueo(wsCal As Worksheet, ByVal colIndex As Long, numOrdenReserva As String, bloqueoArr As Variant, solapeRecolector As Object)
    ' Similar a RegistrarSolape pero para reserva vs bloqueo
    
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If Trim(CStr(numOrdenReserva)) = "-" Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    habStr = ObtenerNumApartamento(bloqueoArr(0))
    
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    ' Crear clave para solape reserva-bloqueo
    If numOrdenReserva <> "-" And IsNumeric(numOrdenReserva) Then
        claveSolape = "AP_" & habStr & "_ORD_" & numOrdenReserva & "_BLOQUEO"
        If Not solapeRecolector.Exists(claveSolape) Then
            solapeRecolector.Add claveSolape, habStr & "|" & numOrdenReserva & "|BLOQUEO|" & CStr(fechaLng) & ","
        Else
            solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
        End If
    End If
End Sub

'=================================================================================
' UTILIDADES PARA TEXTOS - Extraer informaci?n de formatos
'=================================================================================

' Convertir n�mero de fila a nombre de apartamento
Private Function ObtenerNumApartamento(ByVal filaAp As Variant) As String
    On Error Resume Next
    Select Case CLng(filaAp)
        Case 2: ObtenerNumApartamento = "Ap.2"
        Case 4: ObtenerNumApartamento = "Ap.4"
        Case 6: ObtenerNumApartamento = "Ap.5"
        Case 8: ObtenerNumApartamento = "Ap.6"
        Case 10: ObtenerNumApartamento = "Ap.7"
        Case 12: ObtenerNumApartamento = "Ap.8"
        Case 14: ObtenerNumApartamento = "Ap.9"
        Case 16: ObtenerNumApartamento = "Ap.10"
        Case 18: ObtenerNumApartamento = "Ap.12"
        Case Else: ObtenerNumApartamento = "?"
    End Select
End Function

' Extraer N� Orden del texto de reserva
Private Function ObtenerNumOrden(ByVal textoReserva As Variant) As String
    On Error Resume Next
    ObtenerNumOrden = "-"
    
    If IsNull(textoReserva) Then Exit Function
    Dim s As String
    s = CStr(textoReserva)
    If s = "" Or s = "BLOQUEO" Then Exit Function
    
    ' Si no hay salto de l�nea, es solo el N� Orden
    If InStr(s, vbNewLine) = 0 Then
        s = Trim(Replace(s, " C", ""))  ' Quitar "C" de comisi?n
        If IsNumeric(s) Then ObtenerNumOrden = s
        Exit Function
    End If
    
    ' Si hay salto, el N� Orden est� despu�s
    Dim partes() As String
    partes = Split(s, vbNewLine)
    If UBound(partes) < 1 Then Exit Function
    
    Dim lineaNumero As String
    lineaNumero = Trim(partes(1))
    
    ' Quitar "C" si existe
    If Right(lineaNumero, 2) = " C" Then
        lineaNumero = Left(lineaNumero, Len(lineaNumero) - 2)
    End If
    
    ' Extraer primer n�mero
    Dim posibleNum As String
    posibleNum = Trim(Split(lineaNumero, " ")(0))
    If IsNumeric(posibleNum) Then ObtenerNumOrden = posibleNum
End Function

'=================================================================================
' MOSTRAR DETALLE DE SOLAPES - Informe para usuario
'=================================================================================

Private Sub MostrarDetalleSolapes()
    ' OBJETIVO: Mostrar mensaje con detalles de todos los solapes detectados
    
    On Error Resume Next
    If tipoSolape Is Nothing Then Exit Sub
    If tipoSolape.Count = 0 Then
        MsgBox "No hay solapes registrados.", vbInformation, "Sin Solapes"
        Exit Sub
    End If
    
    Dim detalles As String
    detalles = "ATENCION: Solapes Detectados" & vbCrLf & vbCrLf & _
               "Hay reservas que solapan entre si o con bloqueos." & vbCrLf & _
               "Las celdas afectadas aparecen en ROJO." & vbCrLf & vbCrLf & _
               "DETALLE DE SOLAPES DETECTADOS:" & vbCrLf & vbCrLf
    
    Dim k As Variant, info As String, partes() As String
    Dim apStr As String, ord1 As String, ord2 As String, fechasStr As String
    Dim fechasArray() As String, fechasLong() As Long, i As Long, rangoTexto As String
    
    ' Recorrer todos los solapes registrados
    For Each k In tipoSolape.Keys
        info = CStr(tipoSolape(k))
        If InStr(info, "|") = 0 Then GoTo NextSolape
        
        partes = Split(info, "|")
        If UBound(partes) < 3 Then GoTo NextSolape
        
        ' Extraer informaci�n
        apStr = Trim(partes(0))   ' Apartamento
        ord1 = Trim(partes(1))    ' Primer N� Orden
        ord2 = Trim(partes(2))    ' Segundo N� Orden o "BLOQUEO"
        fechasStr = Trim(partes(3))  ' Fechas de solape separadas por coma
        
        ' Limpiar cadena de fechas
        If Right(fechasStr, 1) = "," Then fechasStr = Left(fechasStr, Len(fechasStr) - 1)
        If fechasStr = "" Then GoTo NextSolape
        
        ' Convertir fechas a array
        If InStr(fechasStr, ",") > 0 Then
            fechasArray = Split(fechasStr, ",")
        Else
            ReDim fechasArray(0 To 0)
            fechasArray(0) = fechasStr
        End If
        
        ' Convertir a formato Long
        ReDim fechasLong(LBound(fechasArray) To UBound(fechasArray))
        For i = LBound(fechasArray) To UBound(fechasArray)
            If IsNumeric(Trim(fechasArray(i))) Then
                fechasLong(i) = CLng(Trim(fechasArray(i)))
            End If
        Next i
        
        ' Formatear rango de fechas (agrupar consecutivas)
        rangoTexto = AgruparFechasEnRangos(fechasLong)
        
        ' Agregar al mensaje seg�n tipo de solape
        Dim padding As String
        If Len(rangoTexto) < 18 Then
            padding = Space$(18 - Len(rangoTexto))
        Else
            padding = " "
        End If
        If ord2 = "BLOQUEO" Then
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " " & apStr & vbTab & Chr(187) & " [!] Reserva N" & Chr(186) & " " & ord1 & " con BLOQUEO" & vbCrLf
        Else
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " " & apStr & vbTab & Chr(187) & " [!] Reservas N" & Chr(186) & " " & ord1 & " y " & ord2 & vbCrLf
        End If
        
NextSolape:
    Next k
    
    ' Mostrar mensaje con todos los solapes
    MsgBox detalles, vbExclamation, "Solapes Detectados"
End Sub

Private Function AgruparFechasEnRangos(fechasLong() As Long) As String
    On Error Resume Next
    
    If UBound(fechasLong) < LBound(fechasLong) Then
        AgruparFechasEnRangos = "Sin fechas"
        Exit Function
    End If
    
    Dim fechasList As Collection
    Set fechasList = New Collection
    Dim i As Long
    
    For i = LBound(fechasLong) To UBound(fechasLong)
        If fechasLong(i) > 0 Then
            InsertarFechaOrdenada fechasList, CDate(fechasLong(i))
        End If
    Next i
    
    If fechasList.Count = 0 Then
        AgruparFechasEnRangos = "Sin fechas"
        Exit Function
    End If
    
    Dim resultado As String
    Dim inicioRango As Date, finRango As Date
    Dim primerRango As Boolean: primerRango = True
    
    inicioRango = fechasList(1)
    finRango = fechasList(1)
    
    For i = 2 To fechasList.Count
        If fechasList(i) = finRango + 1 Then
            finRango = fechasList(i)
        Else
            If Not primerRango Then resultado = resultado & ", "
            primerRango = False
            
            If inicioRango = finRango Then
                resultado = resultado & FormatearFechaSolape(inicioRango)
            Else
                resultado = resultado & FormatearFechaSolape(inicioRango) & " al " & FormatearFechaSolape(finRango)
            End If
            
            inicioRango = fechasList(i)
            finRango = fechasList(i)
        End If
    Next i
    
    If Not primerRango Then resultado = resultado & ", "
    
    If inicioRango = finRango Then
        resultado = resultado & FormatearFechaSolape(inicioRango)
    Else
        resultado = resultado & FormatearFechaSolape(inicioRango) & " al " & FormatearFechaSolape(finRango)
    End If
    
    If Year(inicioRango) <> Year(Date) Then
        resultado = resultado & "-" & Year(inicioRango)
    End If
    
    AgruparFechasEnRangos = resultado
End Function

Private Function FormatearFechaSolape(ByVal d As Date) As String
    Dim s As String
    s = Format(d, "d-mmm")
    If Right$(s, 1) = "." Then s = Left$(s, Len(s) - 1)
    
    Dim partes() As String
    partes = Split(s, "-")
    If UBound(partes) = 1 Then
        Dim mes As String
        mes = partes(1)
        If Len(mes) > 0 Then
            mes = UCase$(Left$(mes, 1)) & Mid$(mes, 2)
        End If
        FormatearFechaSolape = partes(0) & "-" & mes
    Else
        FormatearFechaSolape = s
    End If
End Function

' Insertar fecha en colecci?n manteniendo orden
Private Sub InsertarFechaOrdenada(ByRef col As Collection, fecha As Date)
    On Error Resume Next
    
    If col.Count = 0 Then
        col.Add fecha
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To col.Count
        If fecha < col(i) Then
            col.Add fecha, Before:=i  ' Insertar antes si es menor
            Exit Sub
        ElseIf fecha = col(i) Then
            Exit Sub  ' No duplicar
        End If
    Next i
    
    col.Add fecha  ' Agregar al final si es mayor
End Sub

'=================================================================================
' CONTAR RESERVAS EN APARTAMENTOS OCULTOS - Advertencia
'=================================================================================

Private Function ContarReservasEnApartamentosOcultos(wsRes As Worksheet, wsCal As Worksheet, ByRef listaOcultas As String) As Long
    ' OBJETIVO: Contar reservas actuales/futuras asignadas a apartamentos
    '           que est�n ocultos/as en el calendario (filas ocultas)
    '           y recopilar una lista de sus nombres
    
    On Error Resume Next
    
    ContarReservasEnApartamentosOcultos = 0
    
    If wsRes Is Nothing Or wsCal Is Nothing Then Exit Function
    
    Dim lastRow As Long, i As Long
    Dim fechaSalida As Date, hoy As Date
    Dim apartamentos As String, arrAps As Variant
    Dim ap As Variant, filaAp As Long
    Dim contadorOcultas As Long
    
    hoy = Date
    contadorOcultas = 0
    
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    ' Recorrer todas las reservas
    For i = 2 To lastRow
        ' Solo reservas con fecha de salida y apartamento asignado
        If IsDate(wsRes.Cells(i, "M").Value) And Not IsEmpty(wsRes.Cells(i, "S").Value) Then
            fechaSalida = wsRes.Cells(i, "M").Value
            
            ' Solo reservas actuales o futuras (salida >= hoy)
            If fechaSalida >= hoy Then
                apartamentos = Trim(CStr(wsRes.Cells(i, "S").Value))
                
                ' Verificar si hay m�ltiples apartamentos separados por coma
                If InStr(apartamentos, ",") > 0 Then
                    arrAps = Split(apartamentos, ",")
                    Dim resContada As Boolean: resContada = False
                    For Each ap In arrAps
                        filaAp = MapearApartamentoAFila(Trim(CStr(ap)))
                        If filaAp > 0 Then
                            ' Si el apartamento est� oculto en el calendario
                            If wsCal.Rows(filaAp).Hidden Then
                                If Not resContada Then
                                    contadorOcultas = contadorOcultas + 1
                                    resContada = True
                                End If
                                Dim apStr As String
                                apStr = Trim(CStr(ap))
                                If InStr("," & listaOcultas & ",", "," & apStr & ",") = 0 Then
                                    If listaOcultas = "" Then
                                        listaOcultas = apStr
                                    Else
                                        listaOcultas = listaOcultas & ", " & apStr
                                    End If
                                End If
                            End If
                        End If
                    Next ap
                Else
                    ' Reserva simple a un apartamento
                    filaAp = MapearApartamentoAFila(apartamentos)
                    If filaAp > 0 Then
                        If wsCal.Rows(filaAp).Hidden Then
                            contadorOcultas = contadorOcultas + 1
                            Dim apStr2 As String
                            apStr2 = Trim(CStr(apartamentos))
                            If InStr("," & listaOcultas & ",", "," & apStr2 & ",") = 0 Then
                                If listaOcultas = "" Then
                                    listaOcultas = apStr2
                                Else
                                    listaOcultas = listaOcultas & ", " & apStr2
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    ContarReservasEnApartamentosOcultos = contadorOcultas
End Function

' Mapear nombre de apartamento a n�mero de fila en calendario
Private Function MapearApartamentoAFila(ByVal apartamento As String) As Long
    On Error Resume Next
    MapearApartamentoAFila = 0
    
    Select Case Trim(apartamento)
        Case "Ap.2": MapearApartamentoAFila = 2
        Case "Ap.4": MapearApartamentoAFila = 4
        Case "Ap.5": MapearApartamentoAFila = 6
        Case "Ap.6": MapearApartamentoAFila = 8
        Case "Ap.7": MapearApartamentoAFila = 10
        Case "Ap.8": MapearApartamentoAFila = 12
        Case "Ap.9": MapearApartamentoAFila = 14
        Case "Ap.10": MapearApartamentoAFila = 16
        Case "Ap.12": MapearApartamentoAFila = 18
    End Select
End Function



