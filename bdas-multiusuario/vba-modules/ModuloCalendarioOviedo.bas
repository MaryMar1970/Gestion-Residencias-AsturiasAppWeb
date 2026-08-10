Attribute VB_Name = "ModuloCalendarioOviedo"
'Attribute VB_Name = "ModuloCalendarioOviedo" v14.11.3
Option Explicit
Private Const COLOR_SOLAPE_ACEPTADO As Long = 15773696 ' RGB(176, 224, 230) = PowderBlue (azul claro)
'=================================================================================
' Modulo: ModuloCalendarioOviedo
' Version: 17.1 - Sistema de Cache por Fecha de Corte + Correcci?n Limpieza Bloqueos
' Fecha: 2025-01-31
'
' CAMBIOS EN VERSI?N 17.1:
' - Mejoradas funciones de limpieza para manejar correctamente merged cells
' - Agregada funci?n LimpiarBloqueosParcial para limpieza con cach�
' - Modificado flujo en ActualizarCalendarioInterno para llamar a LimpiarBloqueosParcial
' - Agregados comentarios detallados en todo el c�digo
'=================================================================================

' Variable privada para almacenar informaci?n de solapes detectados
Private tipoSolape As Object
Private solapesAceptados As Object

'=================================================================================
' SUB: DefinirTipoSolape
' Inicializa o reinicia el diccionario que almacena los solapes detectados
'=================================================================================
Private Sub DefinirTipoSolape()
    If tipoSolape Is Nothing Then
        Set tipoSolape = CreateObject("Scripting.Dictionary")
    Else
        tipoSolape.RemoveAll
    End If
End Sub

'=================================================================================
' INTERFACES PUBLICAS
' Estas son las funciones que se llaman desde fuera del m�dulo
'=================================================================================

'=================================================================================
' SUB: ActualizarCalendarioCompletoOviedo
' Llamado desde: Bot?n Ribbon "Actualizar Oviedo"
' Prop?sito: Actualiza TODO el calendario (hist?rico completo + futuro)
' Marca el cache como cargado para futuras actualizaciones parciales
'=================================================================================
Public Sub ActualizarCalendarioCompletoOviedo()
    ' Actualizar todo sin restricciones (soloFuturas=False, usarCache=False)
    ActualizarCalendarioInterno False, False
    
    ' Marcar que el hist?rico ya est� cargado en cach�
    GestionCacheCalendarios.MarcarCacheOviedoCargado
    
    ' Informar al usuario en la barra de estado
    Application.StatusBar = "Calendario Oviedo: Historico completo cargado y en cache"
    
    ' Resaltar festivos en el calendario
    Call ResaltarFestivos
End Sub

'=================================================================================
' SUB: ActualizarCalendarioFuturoOviedo
' Llamado desde: Evento Worksheet_Activate (al activar hoja Calendario OVIEDO)
' Prop?sito: Solo actualiza fechas futuras, respetando el hist?rico si hay cach�
' Si no hay cach�, funciona igual que ActualizarCalendarioCompletoOviedo
'=================================================================================
Public Sub ActualizarCalendarioFuturoOviedo()
    ' Actualizar solo futuras si hay cach� disponible
    ' soloFuturas=True: solo procesa reservas/bloqueos futuros
    ' usarCache: True si existe cach�, False si no existe
    ActualizarCalendarioInterno True, GestionCacheCalendarios.CacheOviedoCargado()
End Sub

'=================================================================================
' PROCEDIMIENTO PRINCIPAL
'=================================================================================

'=================================================================================
' SUB: ActualizarCalendarioInterno
' Procedimiento central que coordina toda la actualizaci?n del calendario
'
' Par?metros:
'   - soloFuturas: True = solo actualiza desde hoy-3 d�as en adelante
'                  False = actualiza todo el hist?rico + futuro
'   - usarCache: True = respeta el hist?rico ya cargado, solo actualiza desde fecha de corte
'                False = limpia y recarga todo el calendario completo
'
' Flujo:
' 1. Obtener referencias a hojas necesarias
' 2. Crear mapeo de fechas (diccionario fecha -> columna)
' 3. Determinar fecha y columna de corte si aplica cach�
' 4. Limpiar calendario (completo o parcial seg�n cach�)
' 5. Procesar reservas y bloqueos en memoria
' 6. Aplicar reservas y bloqueos al calendario
' 7. Repintar reservas pagadas en verde
' 8. Mostrar alertas si hay reservas ocultas o solapes
'=================================================================================
Private Sub ActualizarCalendarioInterno(ByVal soloFuturas As Boolean, ByVal usarCache As Boolean)
    ' ============================================================================
    ' DECLARACI?N DE VARIABLES
    ' ============================================================================
    Dim wsRes As Worksheet        ' Hoja "RESIDENCIA OVIEDO" (datos de reservas)
    Dim wsCal As Worksheet        ' Hoja "Calendario OVIEDO" (visual del calendario)
    Dim wsBloq As Worksheet       ' Hoja "Habitaciones Bloqueadas OVIEDO"
    Dim dateMapping As Object     ' Diccionario: fecha -> columna del calendario
    Dim reservasProcessadas As Collection    ' Colecci?n de arrays de reservas procesadas
    Dim bloqueosProcessados As Collection    ' Colecci?n de arrays de bloqueos procesados
    Dim mapaOcupacion As Object   ' Diccionario: "fila_columna" -> texto de ocupaci�n
    Dim huboSolape As Boolean     ' Flag: True si se detectaron solapes
    Dim reservasOcultas As Long   ' Contador de reservas en habitaciones ocultas
    Dim fechaCorte As Date        ' Fecha desde la cual se actualiza (si hay cach�)
    Dim colCorte As Long          ' Columna desde la cual se limpia (si hay cach�)
    Dim estabaProtegida As Boolean  ' Detectar si la hoja del calendario est� protegida
    
    On Error GoTo ErrorHandler
    
    ' ============================================================================
    ' OBTENER REFERENCIAS A HOJAS
    ' ============================================================================
    On Error Resume Next
    Set wsRes = ThisWorkbook.Worksheets("RESIDENCIA OVIEDO")
    Set wsCal = ThisWorkbook.Worksheets("Calendario OVIEDO")
    Set wsBloq = ThisWorkbook.Worksheets("Habitaciones Bloqueadas OVIEDO")
    On Error GoTo ErrorHandler
    
    ' Verificar que todas las hojas existen
    If wsRes Is Nothing Or wsCal Is Nothing Or wsBloq Is Nothing Then
        MsgBox "ERROR: No se encontraron las hojas necesarias para OVIEDO.", vbCritical, "Error"
        Exit Sub
    End If
    
    ' Desproteger la hoja del calendario temporalmente para realizar la limpieza y pintado sin errores
    estabaProtegida = wsCal.ProtectContents
    If estabaProtegida Then wsCal.Unprotect password:=""
    
    ' ============================================================================
    ' DESACTIVAR ACTUALIZACIONES PARA MEJORAR RENDIMIENTO
    ' ============================================================================
    Application.screenUpdating = False     ' No actualizar pantalla mientras se procesa
    Application.enableEvents = False       ' No disparar eventos (evita loops infinitos)
    Application.calculation = xlCalculationManual   ' No recalcular f?rmulas autom�ticamente
    
    ' ============================================================================
    ' CREAR MAPEO DE FECHAS (Diccionario: fecha -> columna)
    ' ============================================================================
    Set dateMapping = CrearMapeoFechas(wsCal)
    If dateMapping Is Nothing Or dateMapping.Count = 0 Then
        MsgBox "ERROR: No se pudo crear mapeo de fechas.", vbExclamation, "Error"
        GoTo CleanExit
    End If
    
    ' ============================================================================
    ' DETERMINAR FECHA DE CORTE Y COLUMNA DE CORTE (si se usa cach�)
    ' ============================================================================
    ' Si usarCache = True, significa que el hist?rico ya est� cargado
    ' Solo necesitamos actualizar desde una fecha espec?fica hacia adelante
    If usarCache Then
        ' Obtener la fecha desde la cual se hizo la �ltima actualizaci?n completa
        fechaCorte = GestionCacheCalendarios.ObtenerFechaCorteOviedo()
        If fechaCorte = 0 Then fechaCorte = Date
        
        ' Restar 3 d�as de margen para asegurar que no perdemos reservas
        fechaCorte = fechaCorte - 3
    Else
        ' Sin cach�: actualizar todo el calendario sin corte
        fechaCorte = 0
    End If
    
    ' Obtener la columna correspondiente a la fecha de corte
    colCorte = 0
    If fechaCorte > 0 Then
        ' Buscar la columna exacta de la fecha de corte
        If dateMapping.Exists(CLng(fechaCorte)) Then
            colCorte = dateMapping(CLng(fechaCorte))
        Else
            ' Si la fecha exacta no existe, buscar la fecha m�s cercana (hasta 10 d�as despu�s)
            Dim fechaBuscar As Date
            For fechaBuscar = fechaCorte To fechaCorte + 10
                If dateMapping.Exists(CLng(fechaBuscar)) Then
                    colCorte = dateMapping(CLng(fechaBuscar))
                    Exit For
                End If
            Next fechaBuscar
        End If
    End If
    
    ' ============================================================================
    ' LIMPIAR CALENDARIO (completo o parcial seg�n cach�)
    ' ============================================================================
    ' CAMBIO v17.1: Ahora se llama tambi?n a LimpiarBloqueosParcial cuando hay cach�
    ' Esto asegura que el texto "BLOQUEADA" se elimine correctamente
    If usarCache And colCorte > 0 Then
        ' MODO PARCIAL (con cach�): Solo limpiar desde colCorte en adelante
        LimpiarCalendarioParcial wsCal, colCorte
        LimpiarBloqueosParcial wsCal, colCorte  ' ? NUEVA LLAMADA v17.1
    Else
        ' MODO COMPLETO (sin cach�): Limpiar todo el calendario
        LimpiarCalendarioRapido wsCal
        LimpiarBloqueos wsCal
    End If
    
    ' ============================================================================
    ' PROCESAR RESERVAS Y BLOQUEOS EN MEMORIA
    ' ============================================================================
    ' Procesar reservas: leer hoja RESIDENCIA OVIEDO y crear arrays de reservas
    Set reservasProcessadas = ProcesarReservasEnMemoriaV13(wsRes, dateMapping, soloFuturas, fechaCorte)
    
    ' Procesar bloqueos: leer hoja Habitaciones Bloqueadas OVIEDO y crear arrays de bloqueos
    Set bloqueosProcessados = ProcesarBloqueosEnMemoriaV13(wsBloq, dateMapping, soloFuturas, fechaCorte)
    
    ' Contar reservas asignadas a habitaciones ocultas (filas ocultas en calendario)
    Dim listaOcultas As String: listaOcultas = ""
    reservasOcultas = ContarReservasEnHabitacionesOcultas(wsRes, wsCal, listaOcultas)
    
    ' ============================================================================
    ' APLICAR RESERVAS Y BLOQUEOS AL CALENDARIO
    ' ============================================================================
    ' Inicializar diccionario de solapes
    DefinirTipoSolape
    
    ' Crear mapa de ocupaci�n: diccionario "fila_columna" -> "texto"
    ' Permite detectar cuando dos reservas intentan ocupar la misma celda (solape)
    Set solapesAceptados = CargarSolapesAceptadosOviedo()
    Set mapaOcupacion = CreateObject("Scripting.Dictionary")
    
    ' Aplicar reservas y bloqueos al calendario visual
    ' Retorna True si se detectaron solapes
    huboSolape = AplicarReservasYBloqueosConTramosV16(wsCal, reservasProcessadas, bloqueosProcessados, mapaOcupacion, tipoSolape)
    
    ' ============================================================================
    ' REPINTAR RESERVAS PAGADAS EN VERDE
    ' ============================================================================
    RepintarReservasPagadas wsRes, wsCal, reservasProcessadas
    
    ' ============================================================================
    ' MOSTRAR ALERTAS AL USUARIO
    ' ============================================================================
    ' Alerta de reservas en habitaciones ocultas
    If reservasOcultas > 0 Then
        Dim mensajeOcultas As String
        If reservasOcultas = 1 Then
            mensajeOcultas = "ATENCION: Hay 1 reserva actual/futura asignada a una habitacion oculta (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Esta reserva NO se muestra en el calendario." & vbCrLf & vbCrLf & _
                           "Para verla, debes mostrar la fila correspondiente en 'Calendario OVIEDO'."
        Else
            mensajeOcultas = "ATENCION: Hay " & reservasOcultas & " reservas actuales/futuras asignadas a habitaciones ocultas (" & listaOcultas & ")." & vbCrLf & vbCrLf & _
                           "Estas reservas NO se muestran en el calendario." & vbCrLf & vbCrLf & _
                           "Para verlas, debes mostrar las filas correspondientes en 'Calendario OVIEDO'."
        End If
        MsgBox mensajeOcultas, vbExclamation, "Reservas en Habitaciones Ocultas"
    End If
    
    ' Alerta de solapes detectados
    If huboSolape Then
        MostrarDetalleSolapes
    End If

CleanExit:
    ' ============================================================================
    ' RESTAURAR CONFIGURACI�N DE EXCEL
    ' ============================================================================
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    Exit Sub

ErrorHandler:
    ' ============================================================================
    ' MANEJO DE ERRORES
    ' ============================================================================
    If estabaProtegida And Not wsCal Is Nothing Then
        wsCal.Protect password:="", UserInterfaceOnly:=True, AllowFormattingCells:=True
    End If
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    MsgBox "ERROR: " & Err.Description, vbCritical, "Error"
End Sub

'=================================================================================
' LIMPIEZA DE CALENDARIO
' CAMBIOS v17.1: Funciones mejoradas para manejar correctamente merged cells
'                y reseteo completo de formato
'=================================================================================

'=================================================================================
' SUB: LimpiarCalendarioRapido
' Limpia TODO el calendario (hist?rico + futuro)
'
' CAMBIOS v17.1:
' - Agregado UnMerge para deshacer celdas combinadas ANTES de limpiar
' - Agregado reseteo de Font.Size y Font.Bold
' - Usa variable rango para mejor rendimiento
'=================================================================================
Private Sub LimpiarCalendarioRapido(wsCal As Worksheet)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    
    Dim ultimaFila As Long, ultimaCol As Long
    Dim rango As Range
    
    ' Encontrar l?mites del calendario
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    ' Validar que hay datos para limpiar
    If ultimaFila < 2 Or ultimaCol < 2 Then Exit Sub
    
    ' Definir rango a limpiar (todo excepto encabezados de fila 1 y columna 1)
    Set rango = wsCal.Range(wsCal.Cells(2, 2), wsCal.Cells(ultimaFila, ultimaCol))
    
    ' CR?TICO: Deshacer combinaciones de celdas ANTES de limpiar
    ' Sin esto, el texto "BLOQUEADA" puede quedar residual
    On Error Resume Next
    rango.UnMerge
    On Error GoTo 0
    
    ' Limpiar contenido y formato
    rango.ClearContents                ' Borra valores y f?rmulas
    rango.Interior.ColorIndex = xlNone ' Borra colores de fondo
    rango.Font.Size = 11               ' Resetea tama?o de fuente a valor por defecto
    rango.Font.Bold = False            ' Resetea negrita a False
    
    Set rango = Nothing
End Sub

'=================================================================================
' SUB: LimpiarCalendarioParcial
' Limpia solo desde la columna indicada en adelante (usado cuando hay cach�)
'
' Par?metros:
'   - colDesde: Columna desde la cual se empieza a limpiar
'
' CAMBIOS v17.1:
' - Agregado UnMerge para deshacer celdas combinadas en rango parcial
' - Agregado reseteo de Font.Size y Font.Bold
' - Usa variable rango para mejor rendimiento
'=================================================================================
Private Sub LimpiarCalendarioParcial(wsCal As Worksheet, ByVal colDesde As Long)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colDesde < 2 Then colDesde = 2  ' Asegurar que no se limpien encabezados
    
    Dim ultimaFila As Long, ultimaCol As Long
    Dim rango As Range
    
    ' Encontrar l?mites del calendario
    ultimaFila = wsCal.Cells(wsCal.Rows.Count, 1).End(xlUp).Row
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    
    ' Validar que hay datos para limpiar
    If ultimaFila < 2 Or ultimaCol < colDesde Then Exit Sub
    
    ' Definir rango parcial a limpiar (desde colDesde hasta el final)
    Set rango = wsCal.Range(wsCal.Cells(2, colDesde), wsCal.Cells(ultimaFila, ultimaCol))
    
    ' CR?TICO: Deshacer combinaciones de celdas ANTES de limpiar
    On Error Resume Next
    rango.UnMerge
    On Error GoTo 0
    
    ' Limpiar contenido y formato
    rango.ClearContents                ' Borra valores y f?rmulas
    rango.Interior.ColorIndex = xlNone ' Borra colores de fondo
    rango.Font.Size = 11               ' Resetea tama?o de fuente
    rango.Font.Bold = False            ' Resetea negrita
    
    Set rango = Nothing
End Sub

'=================================================================================
' SUB: LimpiarBloqueos
' Limpia texto "BLOQUEADA" y color naranja de bloqueos en TODO el calendario
'
' CAMBIOS v17.1:
' - Agregada verificaci?n de MergeCells en cada celda
' - Deshace merge si existe ANTES de limpiar
' - Agregado reseteo de Font.Size y Font.Bold
' - Usa variable celda para claridad
'=================================================================================
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
' CREAR MAPEO DE FECHAS
'=================================================================================

'=================================================================================
' FUNCI�N: CrearMapeoFechas
' Crea un diccionario que mapea cada fecha del calendario a su columna
'
' Retorna: Object (Dictionary) con estructura:
'          Key = CLng(fecha) : fecha como n�mero long
'          Value = columna : n�mero de columna en el calendario
'
' El calendario tiene una estructura de 3 columnas por d?a:
' - Columna 1 (col): Fecha (encabezado)
' - Columna 2 (col+1): Datos adicionales
' - Columna 3 (col+2): Celda principal con nombre/orden
'
' Por eso el Step 3 en el For
'=================================================================================
Private Function CrearMapeoFechas(wsCal As Worksheet) As Object
    On Error Resume Next
    Set CrearMapeoFechas = Nothing
    If wsCal Is Nothing Then Exit Function
    
    Dim dateMapping As Object
    Set dateMapping = CreateObject("Scripting.Dictionary")
    Dim ultimaCol As Long, col As Long
    
    ' Encontrar la �ltima columna con datos
    ultimaCol = wsCal.Cells(1, wsCal.Columns.Count).End(xlToLeft).Column
    If ultimaCol < 2 Then Set CrearMapeoFechas = dateMapping: Exit Function
    
    ' Recorrer las columnas de 3 en 3 (estructura del calendario)
    For col = 2 To ultimaCol Step 3
        ' Verificar que la celda contiene una fecha v�lida
        If IsDate(wsCal.Cells(1, col).Value) Then
            Dim fechaLng As Long
            ' Convertir fecha a Long para usar como key del diccionario
            fechaLng = CLng(wsCal.Cells(1, col).Value)
            
            ' Agregar al diccionario si no existe ya
            If Not dateMapping.Exists(fechaLng) Then dateMapping.Add fechaLng, col
        End If
    Next col
    
    Set CrearMapeoFechas = dateMapping
End Function

'=================================================================================
' PROCESADO EN MEMORIA
' Estas funciones leen las hojas de datos y crean colecciones de arrays en memoria
' para procesar m�s r?pido que leyendo celda por celda
'=================================================================================

'=================================================================================
' FUNCI�N: ProcesarReservasEnMemoriaV13
' Lee la hoja RESIDENCIA OVIEDO y crea una colecci?n de reservas procesadas
'
' Par?metros:
'   - wsRes: Hoja con datos de reservas
'   - dateMapping: Diccionario fecha -> columna
'   - soloFuturas: True = solo procesar reservas desde hoy-3 d�as
'   - fechaCorte: Fecha de corte si hay cach� (0 = sin corte)
'
' Retorna: Collection de arrays, cada array contiene:
'          (0) = fila en calendario
'          (1) = columna check-in
'          (2) = columna check-out
'          (3) = texto a mostrar (nombre + n� orden)
'=================================================================================
Private Function ProcesarReservasEnMemoriaV13(wsRes As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    Dim reservasProcessadas As Collection
    Set reservasProcessadas = New Collection
    On Error Resume Next
    
    ' Validar par?metros
    If wsRes Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    End If
    
    ' Encontrar �ltima fila con datos
    Dim lastRow As Long
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    ' Leer todo el rango de datos de una vez (mucho m�s r?pido que celda por celda)
    ' Columnas A:T = columnas 1 a 20
    Dim arrRes As Variant
    arrRes = wsRes.Range("A2:T" & lastRow).Value
    If Not IsArray(arrRes) Then Set ProcesarReservasEnMemoriaV13 = reservasProcessadas: Exit Function
    
    Dim i As Long, room As String, arrRooms As Variant, r As Variant, reservaArray As Variant
    Dim hoy As Date: hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro seg�n par?metros
    If fechaCorte > 0 Then
        ' Hay cach�: filtrar desde fecha de corte
        fechaFiltro = fechaCorte
    ElseIf soloFuturas Then
        ' No hay cach� pero solo futuras: filtrar desde hoy-3 d�as
        fechaFiltro = hoy - 3
    Else
        ' Sin filtro: procesar todas las reservas
        fechaFiltro = 0
    End If
    
    ' Procesar cada fila del array
    For i = 1 To UBound(arrRes, 1)
        ' Verificar que la fila tiene los datos necesarios:
        ' - Fecha check-in (columna 12 = L)
        ' - Fecha check-out (columna 13 = M)
        ' - Habitaci?n asignada (columna 20 = T)
        If IsDate(arrRes(i, 12)) And IsDate(arrRes(i, 13)) And Not IsEmpty(arrRes(i, 20)) Then
            ' Filtrar por fecha si aplica
            If fechaFiltro > 0 Then
                ' Si el check-out es anterior a la fecha de filtro, saltar esta reserva
                If CDate(arrRes(i, 13)) < fechaFiltro Then GoTo NextI
            End If
            
            ' Obtener habitaci�n(es) asignada(s)
            room = Trim(CStr(arrRes(i, 20)))
            
            ' Una reserva puede tener m�ltiples habitaciones separadas por comas
            If InStr(room, ",") > 0 Then
                ' Procesar cada habitaci�n por separado
                arrRooms = Split(room, ",")
                For Each r In arrRooms
                    If Trim(r) <> "" Then
                        reservaArray = CrearReservaArrayV12(arrRes, i, Trim(r), dateMapping)
                        If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
                    End If
                Next r
            Else
                ' Una sola habitaci�n
                reservaArray = CrearReservaArrayV12(arrRes, i, room, dateMapping)
                If reservaArray(0) > 0 Then reservasProcessadas.Add reservaArray
            End If
        End If
NextI:
    Next i
    Set ProcesarReservasEnMemoriaV13 = reservasProcessadas
End Function

'=================================================================================
' FUNCI�N: ProcesarBloqueosEnMemoriaV13
' Lee la hoja Habitaciones Bloqueadas OVIEDO y crea una colecci?n de bloqueos
'
' Par?metros:
'   - wsBloq: Hoja con bloqueos
'   - dateMapping: Diccionario fecha -> columna
'   - soloFuturas: True = solo procesar bloqueos desde hoy-3 d�as
'   - fechaCorte: Fecha de corte si hay cach� (0 = sin corte)
'
' Retorna: Collection de arrays, cada array contiene:
'          (0) = fila en calendario
'          (1) = columna inicio bloqueo
'          (2) = columna fin bloqueo
'          (3) = habitaci�n bloqueada
'=================================================================================
Private Function ProcesarBloqueosEnMemoriaV13(wsBloq As Worksheet, dateMapping As Object, ByVal soloFuturas As Boolean, ByVal fechaCorte As Date) As Collection
    Dim bloqueosProcessados As Collection
    Set bloqueosProcessados = New Collection
    On Error Resume Next
    
    ' Validar par?metros
    If wsBloq Is Nothing Or dateMapping Is Nothing Then
        Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    End If
    
    ' Encontrar �ltima fila con datos
    Dim lastRow As Long
    lastRow = wsBloq.Cells(wsBloq.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    ' Leer todo el rango de datos de una vez
    ' Columnas A:C = Habitaci?n, Fecha Inicio, Fecha Fin
    Dim arrBloq As Variant
    arrBloq = wsBloq.Range("A2:C" & lastRow).Value
    If Not IsArray(arrBloq) Then Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados: Exit Function
    
    Dim i As Long, bloqueoArray As Variant, hoy As Date
    hoy = Date
    Dim fechaFiltro As Date
    
    ' Determinar fecha de filtro
    If fechaCorte > 0 Then
        fechaFiltro = fechaCorte
    ElseIf soloFuturas Then
        fechaFiltro = hoy - 3
    Else
        fechaFiltro = 0
    End If
    
    ' Procesar cada fila del array
    For i = 1 To UBound(arrBloq, 1)
        ' Verificar que la fila tiene los datos necesarios
        If arrBloq(i, 1) <> "" And IsDate(arrBloq(i, 2)) And IsDate(arrBloq(i, 3)) Then
            ' Filtrar por fecha si aplica
            If fechaFiltro > 0 Then
                If CDate(arrBloq(i, 3)) < fechaFiltro Then GoTo NextJ
            End If
            
            bloqueoArray = CrearBloqueoArrayV12(arrBloq, i, dateMapping)
            If bloqueoArray(0) > 0 Then bloqueosProcessados.Add bloqueoArray
        End If
NextJ:
    Next i
    Set ProcesarBloqueosEnMemoriaV13 = bloqueosProcessados
End Function

'=================================================================================
' CREAR ARRAYS DE RESERVA Y BLOQUEO
' Estas funciones convierten los datos crudos en arrays estructurados
'=================================================================================

'=================================================================================
' FUNCI�N: CrearReservaArrayV12
' Convierte una fila de la hoja RESIDENCIA OVIEDO en un array estructurado
'
' Par?metros:
'   - arrRes: Array con todos los datos de reservas
'   - i: ?ndice de la fila a procesar
'   - room: C?digo de habitaci�n (ej: "1", "5", "12")
'   - dateMapping: Diccionario fecha -> columna
'
' Retorna: Array(0 To 3):
'   (0) = Fila en calendario (0 si habitaci�n inv?lida o oculta)
'   (1) = Columna de check-in
'   (2) = Columna de check-out
'   (3) = Texto a mostrar: "Nombre (pax)\nN? Orden[ C]"
'=================================================================================
Private Function CrearReservaArrayV12(arrRes As Variant, i As Long, ByVal room As String, dateMapping As Object) As Variant
    Dim reservaArray(0 To 3) As Variant
    reservaArray(0) = 0: reservaArray(1) = 0: reservaArray(2) = 0: reservaArray(3) = ""
    On Error Resume Next
    
    ' Mapear c�digo de habitaci�n a fila en calendario
    ' La estructura del calendario es: cada habitaci�n ocupa 2 filas (par/impar)
    Select Case room
        Case "1": reservaArray(0) = 2
        Case "2": reservaArray(0) = 4
        Case "3": reservaArray(0) = 6
        Case "4": reservaArray(0) = 8
        Case "5": reservaArray(0) = 10
        Case "6": reservaArray(0) = 12
        Case "7": reservaArray(0) = 14
        Case "8": reservaArray(0) = 16
        Case "9": reservaArray(0) = 18
        Case "10": reservaArray(0) = 20
        Case "11": reservaArray(0) = 22
        Case "12": reservaArray(0) = 24
        Case "13": reservaArray(0) = 26
        Case "14": reservaArray(0) = 28
        Case "15": reservaArray(0) = 30
        Case Else
            ' Habitaci?n no reconocida
            CrearReservaArrayV12 = reservaArray
            Exit Function
    End Select
    
    ' Verificar si la fila de esta habitaci�n est� oculta en el calendario
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario OVIEDO")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(reservaArray(0)).Hidden Then
            ' Habitaci?n oculta: no procesar esta reserva
            reservaArray(0) = 0
            CrearReservaArrayV12 = reservaArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ' Verificar fechas v�lidas
    If Not IsDate(arrRes(i, 12)) Or Not IsDate(arrRes(i, 13)) Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    ' Obtener fechas y buscar sus columnas en el calendario
    Dim checkin As Date, checkout As Date
    checkin = CDate(arrRes(i, 12))
    checkout = CDate(arrRes(i, 13))
    
    If dateMapping.Exists(CLng(checkin)) Then reservaArray(1) = dateMapping(CLng(checkin))
    If dateMapping.Exists(CLng(checkout)) Then reservaArray(2) = dateMapping(CLng(checkout))
    
    ' Si alguna fecha no est� en el calendario, descartar reserva
    If reservaArray(1) = 0 Or reservaArray(2) = 0 Then
        reservaArray(0) = 0: CrearReservaArrayV12 = reservaArray: Exit Function
    End If
    
    ' Construir texto a mostrar en el calendario
    Dim fullName As String, displayName As String
    fullName = Trim(CStr(arrRes(i, 11)))  ' Columna 11 = K = Nombre completo
    
    If fullName = "" Then
        ' Sin nombre: solo mostrar n�mero de orden
        reservaArray(3) = arrRes(i, 1)
    Else
        ' Leer N� pax de columna O (posici?n 15 en el array)
        Dim pax As String
        pax = Trim(CStr(arrRes(i, 15)))
        
        ' Con nombre: truncar si es muy largo (m�ximo 15 caracteres)
        ' para dejar espacio al "(N pax)"
        If Len(fullName) <= 15 Then
            displayName = fullName & " (" & pax & ")"
        Else
            displayName = Left(fullName, 15) & ". (" & pax & ")"
        End If
        ' Formato: "Nombre (N pax)\nN? Orden"
        reservaArray(3) = displayName & vbNewLine & arrRes(i, 1)
    End If
    
    ' Agregar "C" si es una reserva con comisi?n
    ' Columna 4 = D = Tipo de reserva
    If arrRes(i, 4) = "Comisi�n NO indem." Or arrRes(i, 4) = "Destino" Or arrRes(i, 4) = "Comisi�n" Then
        reservaArray(3) = reservaArray(3) & " C"
    End If
    
    CrearReservaArrayV12 = reservaArray
End Function

'=================================================================================
' FUNCI�N: CrearBloqueoArrayV12
' Convierte una fila de la hoja Habitaciones Bloqueadas OVIEDO en un array
'
' Par?metros:
'   - arrBloq: Array con todos los datos de bloqueos
'   - i: ?ndice de la fila a procesar
'   - dateMapping: Diccionario fecha -> columna
'
' Retorna: Array(0 To 3):
'   (0) = Fila en calendario (0 si habitaci�n inv?lida o oculta)
'   (1) = Columna inicio del bloqueo
'   (2) = Columna fin del bloqueo
'   (3) = C?digo de habitaci�n
'=================================================================================
Private Function CrearBloqueoArrayV12(arrBloq As Variant, i As Long, dateMapping As Object) As Variant
    Dim bloqueoArray(0 To 3) As Variant
    bloqueoArray(0) = 0: bloqueoArray(1) = 0: bloqueoArray(2) = 0
    bloqueoArray(3) = CStr(arrBloq(i, 1))  ' Habitaci?n
    On Error Resume Next
    
    ' Mapear habitaci�n a fila (igual que en reservas)
    Select Case bloqueoArray(3)
        Case "1": bloqueoArray(0) = 2
        Case "2": bloqueoArray(0) = 4
        Case "3": bloqueoArray(0) = 6
        Case "4": bloqueoArray(0) = 8
        Case "5": bloqueoArray(0) = 10
        Case "6": bloqueoArray(0) = 12
        Case "7": bloqueoArray(0) = 14
        Case "8": bloqueoArray(0) = 16
        Case "9": bloqueoArray(0) = 18
        Case "10": bloqueoArray(0) = 20
        Case "11": bloqueoArray(0) = 22
        Case "12": bloqueoArray(0) = 24
        Case "13": bloqueoArray(0) = 26
        Case "14": bloqueoArray(0) = 28
        Case "15": bloqueoArray(0) = 30
        Case Else
            CrearBloqueoArrayV12 = bloqueoArray
            Exit Function
    End Select
    
    ' Verificar si la fila est� oculta
    On Error Resume Next
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Worksheets("Calendario OVIEDO")
    If Not wsCal Is Nothing Then
        If wsCal.Rows(bloqueoArray(0)).Hidden Then
            bloqueoArray(0) = 0
            CrearBloqueoArrayV12 = bloqueoArray
            Exit Function
        End If
    End If
    On Error GoTo 0
    
    ' Verificar fechas v�lidas
    If Not IsDate(arrBloq(i, 2)) Or Not IsDate(arrBloq(i, 3)) Then
        bloqueoArray(0) = 0: CrearBloqueoArrayV12 = bloqueoArray: Exit Function
    End If
    
    ' Obtener fechas y buscar columnas
    Dim dIni As Date, dFin As Date
    dIni = CDate(arrBloq(i, 2))
    dFin = CDate(arrBloq(i, 3))
    
    If dateMapping.Exists(CLng(dIni)) Then bloqueoArray(1) = dateMapping(CLng(dIni))
    If dateMapping.Exists(CLng(dFin)) Then bloqueoArray(2) = dateMapping(CLng(dFin))
    If bloqueoArray(1) = 0 Or bloqueoArray(2) = 0 Then bloqueoArray(0) = 0
    
    CrearBloqueoArrayV12 = bloqueoArray
End Function

'=================================================================================
' APLICAR RESERVAS Y BLOQUEOS
'=================================================================================

'=================================================================================
' FUNCI�N: AplicarReservasYBloqueosConTramosV16
' Aplica las reservas y bloqueos procesados al calendario visual
'
' L?GICA DE TRAMOS (estructura de 3 columnas por d?a):
' - D?a Check-IN: Solo TRAMO 3 (columna +2) en AMARILLO con texto
' - D?as Intermedios: TODOS los tramos (0, 1, 2) en AMARILLO sin texto
' - D?a Check-OUT: Solo TRAMO 1 (columna +0) en AMARILLO sin texto
' - Bloqueos: TODOS los tramos de TODOS los d�as en NARANJA
'
' Par?metros:
'   - wsCal: Hoja del calendario
'   - reservas: Colecci?n de arrays de reservas
'   - bloqueos: Colecci?n de arrays de bloqueos
'   - mapaOcupacion: Diccionario para detectar solapes
'   - solapeRecolector: Diccionario para registrar detalles de solapes
'
' Retorna: Boolean - True si se detectaron solapes
'=================================================================================
Private Function AplicarReservasYBloqueosConTramosV16(wsCal As Worksheet, reservas As Collection, bloqueos As Collection, mapaOcupacion As Object, solapeRecolector As Object) As Boolean
    Dim huboSolape As Boolean: huboSolape = False
    Dim i As Long, reservaArray As Variant, col As Long, clave As String
    Dim tramo As Long, bloqueoArray As Variant
    Dim colCheckin As Long, colCheckout As Long
    
    On Error Resume Next

    ' ============================================================================
    ' APLICAR RESERVAS
    ' ============================================================================
    For i = 1 To reservas.Count
        reservaArray = reservas(i)
        
        colCheckin = reservaArray(1)   ' Columna de check-in
        colCheckout = reservaArray(2)  ' Columna de check-out
        
        ' ------------------------------------------------------------------------
        ' D?A DE CHECK-IN - Solo TRAMO 3 (columna +2)
        ' Esta es la celda principal que muestra el nombre y n�mero de orden
        ' ------------------------------------------------------------------------
        clave = reservaArray(0) & "_" & (colCheckin + 2)
        With wsCal.Cells(reservaArray(0), colCheckin + 2)
            ' Deshacer merge si existe (puede haber quedado de actualizaciones previas)
            If .MergeCells Then .MergeCells = False
            
            ' Verificar si esta celda ya est� ocupada (SOLAPE)
            If mapaOcupacion.Exists(clave) Then
                If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), colCheckin + 2) Then
                    .Interior.color = COLOR_SOLAPE_ACEPTADO
                Else
                    .Interior.color = vbRed
                    huboSolape = True
                    RegistrarSolape wsCal, colCheckin + 2, reservaArray, mapaOcupacion(clave), solapeRecolector
                End If
            Else
                .Interior.color = vbYellow
                mapaOcupacion(clave) = reservaArray(3)
            End If
            
            ' Aplicar texto y formato
            .Value = reservaArray(3)  ' Nombre + N� Orden
            .Font.Size = 15           ' Tama?o del nombre
            
            ' Hacer el n�mero de orden m�s grande y en negrita
            If InStr(.Value, vbNewLine) > 0 Then
                Dim lineBreakPos As Long
                lineBreakPos = InStr(.Value, vbNewLine)
                .Characters(Start:=lineBreakPos + 1, Length:=Len(.Value) - lineBreakPos).Font.Size = 36
                .Font.Bold = True
            End If
        End With
        
        ' ------------------------------------------------------------------------
        ' D?AS INTERMEDIOS - TODOS los tramos (0, 1, 2)
        ' Solo se pintan en amarillo, sin texto
        ' ------------------------------------------------------------------------
        If colCheckout > colCheckin + 3 Then
            For col = colCheckin + 3 To colCheckout - 3 Step 3
                For tramo = 0 To 2
                    clave = reservaArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(reservaArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        
                        ' Verificar solape
                        If mapaOcupacion.Exists(clave) Then
                            If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), col + tramo) Then
                                .Interior.color = COLOR_SOLAPE_ACEPTADO
                            Else
                                .Interior.color = vbRed
                                huboSolape = True
                                RegistrarSolape wsCal, col + tramo, reservaArray, mapaOcupacion(clave), solapeRecolector
                            End If
                        Else
                            .Interior.color = vbYellow
                            mapaOcupacion(clave) = reservaArray(3)
                        End If
                    End With
                Next tramo
            Next col
        End If
        
        ' ------------------------------------------------------------------------
        ' D?A DE CHECK-OUT - Solo TRAMO 1 (columna +0)
        ' Solo se pinta en amarillo, sin texto
        ' ------------------------------------------------------------------------
        If colCheckout > colCheckin Then
            clave = reservaArray(0) & "_" & colCheckout
            With wsCal.Cells(reservaArray(0), colCheckout)
                If .MergeCells Then .MergeCells = False
                
                ' Verificar solape
                If mapaOcupacion.Exists(clave) Then
                    If EsSolapeGrupoAceptadoEnCelda(wsCal, reservaArray(0), colCheckout) Then
                        .Interior.color = COLOR_SOLAPE_ACEPTADO
                    Else
                        .Interior.color = vbRed
                        huboSolape = True
                        RegistrarSolape wsCal, colCheckout, reservaArray, mapaOcupacion(clave), solapeRecolector
                    End If
                Else
                    .Interior.color = vbYellow
                    mapaOcupacion(clave) = reservaArray(3)
                End If
            End With
        End If
        
    Next i

    ' APLICAR BLOQUEOS: pintado identico a reservas (intervalo semiabierto [inicio, fin))
    ' - Dia inicio: solo tramo checkin (col+2), con texto "BLOQUEADA"
    ' - Dias intermedios: los 3 tramos, con texto "BLOQUEADA" en tramo 2
    ' - Dia fin: solo tramo checkout (col+0), sin texto
    For i = 1 To bloqueos.Count
        bloqueoArray = bloqueos(i)
        
        ' DIA DE INICIO DEL BLOQUEO - solo tramo checkin (col+2)
        clave = bloqueoArray(0) & "_" & (bloqueoArray(1) + 2)
        With wsCal.Cells(bloqueoArray(0), bloqueoArray(1) + 2)
            If .MergeCells Then .MergeCells = False
            If mapaOcupacion.Exists(clave) Then
                If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), bloqueoArray(1) + 2) Then
                    .Interior.color = COLOR_SOLAPE_ACEPTADO
                Else
                    .Interior.color = vbRed
                    huboSolape = True
                    Dim numOrdenSolape As String
                    numOrdenSolape = "-"
                    If mapaOcupacion(clave) <> "BLOQUEO" Then
                        numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                    End If
                    RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(1) + 2, numOrdenSolape, bloqueoArray, solapeRecolector
                End If
            Else
                .Interior.color = RGB(255, 192, 0)
                mapaOcupacion(clave) = "BLOQUEO"
            End If
            .Value = "BLOQUEADA"
            .Font.Size = 22
            .Font.Bold = True
        End With
        
        ' DIAS INTERMEDIOS DEL BLOQUEO - los 3 tramos
        If bloqueoArray(2) > bloqueoArray(1) + 3 Then
            For col = bloqueoArray(1) + 3 To bloqueoArray(2) - 3 Step 3
                For tramo = 0 To 2
                    clave = bloqueoArray(0) & "_" & (col + tramo)
                    With wsCal.Cells(bloqueoArray(0), col + tramo)
                        If .MergeCells Then .MergeCells = False
                        If mapaOcupacion.Exists(clave) Then
                            If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), col + tramo) Then
                                .Interior.color = COLOR_SOLAPE_ACEPTADO
                            Else
                                .Interior.color = vbRed
                                huboSolape = True
                                numOrdenSolape = "-"
                                If mapaOcupacion(clave) <> "BLOQUEO" Then
                                    numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                                End If
                                RegistrarSolapeReservaBloqueo wsCal, col + tramo, numOrdenSolape, bloqueoArray, solapeRecolector
                            End If
                        Else
                            .Interior.color = RGB(255, 192, 0)
                            mapaOcupacion(clave) = "BLOQUEO"
                        End If
                    End With
                Next tramo
                With wsCal.Cells(bloqueoArray(0), col + 2)
                    .Value = "BLOQUEADA"
                    .Font.Size = 22
                    .Font.Bold = True
                End With
            Next col
        End If
        
        ' DIA DE FIN DEL BLOQUEO - solo tramo checkout (col+0), si es distinto al inicio
        If bloqueoArray(2) > bloqueoArray(1) Then
            clave = bloqueoArray(0) & "_" & bloqueoArray(2)
            With wsCal.Cells(bloqueoArray(0), bloqueoArray(2))
                If .MergeCells Then .MergeCells = False
                If mapaOcupacion.Exists(clave) Then
                    If EsSolapeGrupoAceptadoEnCelda(wsCal, bloqueoArray(0), bloqueoArray(2)) Then
                        .Interior.color = COLOR_SOLAPE_ACEPTADO
                    Else
                        .Interior.color = vbRed
                        huboSolape = True
                        numOrdenSolape = "-"
                        If mapaOcupacion(clave) <> "BLOQUEO" Then
                            numOrdenSolape = ObtenerNumOrden(mapaOcupacion(clave))
                        End If
                        RegistrarSolapeReservaBloqueo wsCal, bloqueoArray(2), numOrdenSolape, bloqueoArray, solapeRecolector
                    End If
                Else
                    .Interior.color = RGB(255, 192, 0)
                    mapaOcupacion(clave) = "BLOQUEO"
                End If
            End With
        End If
    Next i
    
    AplicarReservasYBloqueosConTramosV16 = huboSolape
End Function

'=================================================================================
' REPINTAR RESERVAS PAGADAS
'=================================================================================

'=================================================================================
' SUB: RepintarReservasPagadas
' Cambia el color de las reservas pagadas de AMARILLO a VERDE
' Color verde: RGB(198, 239, 206)
'
' Busca el campo "Pagado" (columna 28 = AB) en la hoja RESIDENCIA OVIEDO
' Si dice "SI", repinta toda la reserva en verde
'=================================================================================
Private Sub RepintarReservasPagadas(wsRes As Worksheet, wsCal As Worksheet, reservasProcessadas As Collection)
    On Error Resume Next
    
    If reservasProcessadas Is Nothing Then Exit Sub
    If reservasProcessadas.Count = 0 Then Exit Sub
    
    Dim i As Long, pagado As String, filaRes As Long, reservaArray As Variant
    Dim col As Long, tramo As Long, colCheckin As Long, colCheckout As Long
    
    For i = 1 To reservasProcessadas.Count
        reservaArray = reservasProcessadas(i)
        Dim partes() As String
        
        ' Extraer n�mero de orden del texto de la reserva
        If InStr(reservaArray(3), vbNewLine) > 0 Then
            partes = Split(reservaArray(3), vbNewLine)
            If UBound(partes) > 0 Then
                Dim numOrden As Variant
                numOrden = Trim(Split(partes(1), " ")(0))
                
                ' Buscar esta reserva en la hoja RESIDENCIA OVIEDO
                filaRes = BuscarEnColumna(wsRes, 1, numOrden)
                If filaRes > 0 Then
                    ' Verificar si est� pagada (columna 28 = AB)
                    pagado = UCase(wsRes.Cells(filaRes, 28).Value)
                    If MarcarSiPagadosEnResidencia.EsEstadoPagado(pagado) Then
                        ' Repintar en VERDE toda la reserva
                        colCheckin = reservaArray(1)
                        colCheckout = reservaArray(2)
                        
                        ' D?a check-in (tramo 3)
                        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = RGB(198, 239, 206)
                        
                        ' D?as intermedios (todos los tramos)
                        If colCheckout > colCheckin + 3 Then
                            For col = colCheckin + 3 To colCheckout - 3 Step 3
                                For tramo = 0 To 2
                                    wsCal.Cells(reservaArray(0), col + tramo).Interior.color = RGB(198, 239, 206)
                                Next tramo
                            Next col
                        End If
                        
                        ' D?a check-out (tramo 1)
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
' BUSQUEDAS Y UTILIDADES
'=================================================================================

'=================================================================================
' FUNCI�N: BuscarEnColumna
' Busca un valor en una columna espec?fica de una hoja
'
' Retorna: N�mero de fila donde se encontr� el valor (0 si no se encontr�)
'=================================================================================
Private Function BuscarEnColumna(ws As Worksheet, numCol As Long, valor As Variant) As Long
    On Error Resume Next
    BuscarEnColumna = 0
    If ws Is Nothing Then Exit Function
    If numCol < 1 Or numCol > ws.Columns.Count Then Exit Function
    
    Dim ultimaFila As Long, i As Long
    ultimaFila = ws.Cells(ws.Rows.Count, numCol).End(xlUp).Row
    If ultimaFila < 2 Then Exit Function
    
    ' B?squeda lineal (podr?a optimizarse con Find, pero esto es m�s seguro)
    For i = 2 To ultimaFila
        If Trim(CStr(ws.Cells(i, numCol).Value)) = Trim(CStr(valor)) Then
            BuscarEnColumna = i
            Exit Function
        End If
    Next i
End Function

'=================================================================================
' REGISTRO DE SOLAPES
' Estas funciones registran los detalles de los solapes para mostrarlos al usuario
'=================================================================================

'=================================================================================
' SUB: RegistrarSolape
' Registra un solape entre dos reservas
' Almacena: habitaci�n, n�meros de orden, y fechas afectadas
'=================================================================================
Private Sub RegistrarSolape(wsCal As Worksheet, ByVal colIndex As Long, arr1 As Variant, arr2 As Variant, solapeRecolector As Object, Optional esBloqueo As Boolean = False)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, numOrdenReserva As String, numOrdenReserva2 As String
    Dim claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    ' Calcular la primera columna del d?a (cada d?a tiene 3 columnas)
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    ' Obtener fecha de la celda
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    ' Extraer datos del solape
    habStr = ObtenerNumHab(arr1(0))
    
    ' --- FILTRO: SOLAPE ACEPTADO POR GRUPO (HAB + RANGO) ---
    Dim claveGrupo As String
    claveGrupo = ClaveSolapeGrupoOviedo(habStr)
    
    If Not solapesAceptados Is Nothing Then
        If SolapeAceptadoOviedo(solapesAceptados, claveGrupo, CDate(fechaLng)) Then
            Exit Sub
        End If
    End If
    
    numOrdenReserva = ObtenerNumOrden(arr1(3))
    numOrdenReserva2 = ObtenerNumOrden(arr2)
    
    ' Validar datos
    If numOrdenReserva = "-" Or numOrdenReserva2 = "-" Then Exit Sub
    If Not IsNumeric(numOrdenReserva) Or Not IsNumeric(numOrdenReserva2) Then Exit Sub
    
    ' Ordenar n�meros de orden (el menor primero)
    Dim ord1 As Long, ord2 As Long
    ord1 = CLng(numOrdenReserva)
    ord2 = CLng(numOrdenReserva2)
    
    If ord1 > ord2 Then
        Dim temp As Long
        temp = ord1: ord1 = ord2: ord2 = temp
    End If
    
    ' Crear clave ?nica para este solape
    claveSolape = "HAB_" & habStr & "_ORD_" & CStr(ord1) & "_" & CStr(ord2)
    
    ' Registrar o actualizar solape
    If Not solapeRecolector.Exists(claveSolape) Then
        ' Nuevo solape: crear entrada
        solapeRecolector.Add claveSolape, habStr & "|" & CStr(ord1) & "|" & CStr(ord2) & "|" & CStr(fechaLng) & ","
    Else
        ' Solape existente: agregar fecha
        solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
    End If
End Sub

'=================================================================================
' SUB: RegistrarSolapeReservaBloqueo
' Registra un solape entre una reserva y un bloqueo
'=================================================================================
Private Sub RegistrarSolapeReservaBloqueo(wsCal As Worksheet, ByVal colIndex As Long, numOrdenReserva As String, bloqueoArr As Variant, solapeRecolector As Object)
    On Error Resume Next
    If wsCal Is Nothing Then Exit Sub
    If Trim(CStr(numOrdenReserva)) = "-" Then Exit Sub
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Sub
    
    Dim habStr As String, claveSolape As String, fechaLng As Long
    Dim dayFirstCol As Long
    Dim fechaCelda As Variant
    
    habStr = ObtenerNumHab(bloqueoArr(0))
    
    ' Calcular primera columna del d?a
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2
    
    ' Obtener fecha
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Sub
    fechaLng = CLng(CDate(fechaCelda))
    
    ' Registrar solape con bloqueo
    If numOrdenReserva <> "-" And IsNumeric(numOrdenReserva) Then
        claveSolape = "HAB_" & habStr & "_ORD_" & numOrdenReserva & "_BLOQUEO"
        If Not solapeRecolector.Exists(claveSolape) Then
            solapeRecolector.Add claveSolape, habStr & "|" & numOrdenReserva & "|BLOQUEO|" & CStr(fechaLng) & ","
        Else
            solapeRecolector(claveSolape) = solapeRecolector(claveSolape) & CStr(fechaLng) & ","
        End If
    End If
End Sub

'=================================================================================
' UTILIDADES PARA TEXTOS
'=================================================================================

'=================================================================================
' FUNCI�N: ObtenerNumHab
' Convierte n�mero de fila del calendario a c�digo de habitaci�n
'=================================================================================
Private Function ObtenerNumHab(ByVal filaHab As Variant) As String
    On Error Resume Next
    Select Case CLng(filaHab)
        Case 2: ObtenerNumHab = "1"
        Case 4: ObtenerNumHab = "2"
        Case 6: ObtenerNumHab = "3"
        Case 8: ObtenerNumHab = "4"
        Case 10: ObtenerNumHab = "5"
        Case 12: ObtenerNumHab = "6"
        Case 14: ObtenerNumHab = "7"
        Case 16: ObtenerNumHab = "8"
        Case 18: ObtenerNumHab = "9"
        Case 20: ObtenerNumHab = "10"
        Case 22: ObtenerNumHab = "11"
        Case 24: ObtenerNumHab = "12"
        Case 26: ObtenerNumHab = "13"
        Case 28: ObtenerNumHab = "14"
        Case 30: ObtenerNumHab = "15"
        Case Else: ObtenerNumHab = "?"
    End Select
End Function

'=================================================================================
' FUNCI�N: ObtenerNumOrden
' Extrae el n�mero de orden del texto de una reserva
' Formato esperado: "Nombre\nN?[ C]" o solo "N�[ C]"
'=================================================================================
Private Function ObtenerNumOrden(ByVal textoReserva As Variant) As String
    On Error Resume Next
    ObtenerNumOrden = "-"
    
    If IsNull(textoReserva) Then Exit Function
    Dim s As String
    s = CStr(textoReserva)
    If s = "" Or s = "BLOQUEO" Then Exit Function
    
    ' Si no hay salto de l�nea, el texto es solo el n�mero de orden
    If InStr(s, vbNewLine) = 0 Then
        s = Trim(Replace(s, " C", ""))  ' Quitar " C" si existe
        If IsNumeric(s) Then ObtenerNumOrden = s
        Exit Function
    End If
    
    ' Hay salto de l�nea: formato "Nombre\nN?[ C]"
    Dim partes() As String
    partes = Split(s, vbNewLine)
    If UBound(partes) < 1 Then Exit Function
    
    Dim lineaNumero As String
    lineaNumero = Trim(partes(1))  ' Segunda l�nea = n�mero de orden
    
    ' Quitar " C" si existe
    If Right(lineaNumero, 2) = " C" Then
        lineaNumero = Left(lineaNumero, Len(lineaNumero) - 2)
    End If
    
    ' Extraer primer elemento (el n�mero)
    Dim posibleNum As String
    posibleNum = Trim(Split(lineaNumero, " ")(0))
    If IsNumeric(posibleNum) Then ObtenerNumOrden = posibleNum
End Function

'=================================================================================
' MOSTRAR DETALLE DE SOLAPES
'=================================================================================

'=================================================================================
' SUB: MostrarDetalleSolapes
' Muestra un MessageBox con el detalle de todos los solapes detectados
' Agrupa fechas consecutivas en rangos para facilitar lectura
'=================================================================================
Private Sub MostrarDetalleSolapes()
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
    Dim habStr As String, ord1 As String, ord2 As String, fechasStr As String
    Dim fechasArray() As String, fechasLong() As Long, i As Long, rangoTexto As String
    
    For Each k In tipoSolape.Keys
        info = CStr(tipoSolape(k))
        If InStr(info, "|") = 0 Then GoTo NextSolape
        
        partes = Split(info, "|")
        If UBound(partes) < 3 Then GoTo NextSolape
        
        habStr = Trim(partes(0))
        ord1 = Trim(partes(1))
        ord2 = Trim(partes(2))
        fechasStr = Trim(partes(3))
        
        If Right(fechasStr, 1) = "," Then fechasStr = Left(fechasStr, Len(fechasStr) - 1)
        If fechasStr = "" Then GoTo NextSolape
        
        If InStr(fechasStr, ",") > 0 Then
            fechasArray = Split(fechasStr, ",")
        Else
            ReDim fechasArray(0 To 0)
            fechasArray(0) = fechasStr
        End If
        
        ReDim fechasLong(LBound(fechasArray) To UBound(fechasArray))
        For i = LBound(fechasArray) To UBound(fechasArray)
            If IsNumeric(Trim(fechasArray(i))) Then
                fechasLong(i) = CLng(Trim(fechasArray(i)))
            End If
        Next i
        
        rangoTexto = AgruparFechasEnRangos(fechasLong)
        
        Dim padding As String
        If Len(rangoTexto) < 18 Then
            padding = Space$(18 - Len(rangoTexto))
        Else
            padding = " "
        End If
        If ord2 = "BLOQUEO" Then
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " Hab. " & habStr & vbTab & Chr(187) & " [!] Reserva N" & Chr(186) & " " & ord1 & " con BLOQUEO" & vbCrLf
        Else
            detalles = detalles & Chr(149) & " " & rangoTexto & padding & Chr(187) & " Hab. " & habStr & vbTab & Chr(187) & " [!] Reservas N" & Chr(186) & " " & ord1 & " y " & ord2 & vbCrLf
        End If
        
NextSolape:
    Next k
    
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

'=================================================================================
' SUB: InsertarFechaOrdenada
' Inserta una fecha en una colecci?n manteni?ndola ordenada
' Evita duplicados
'=================================================================================
Private Sub InsertarFechaOrdenada(ByRef col As Collection, fecha As Date)
    On Error Resume Next
    
    If col.Count = 0 Then
        col.Add fecha
        Exit Sub
    End If
    
    Dim i As Long
    For i = 1 To col.Count
        If fecha < col(i) Then
            ' Insertar antes de esta posici?n
            col.Add fecha, Before:=i
            Exit Sub
        ElseIf fecha = col(i) Then
            ' Ya existe: no agregar duplicado
            Exit Sub
        End If
    Next i
    
    ' Si llegamos aqu?, la fecha es mayor que todas: agregar al final
    col.Add fecha
End Sub

'=================================================================================
' CONTAR RESERVAS EN HABITACIONES OCULTAS
'=================================================================================

'=================================================================================
' FUNCI�N: ContarReservasEnHabitacionesOcultas
' Cuenta cu�ntas reservas actuales/futuras est�n en habitaciones ocultas
' Una habitaci�n oculta es aquella cuya fila est� oculta en el calendario
'
' Retorna: Cantidad de reservas en habitaciones ocultas
'=================================================================================
Private Function ContarReservasEnHabitacionesOcultas(wsRes As Worksheet, wsCal As Worksheet, ByRef listaOcultas As String) As Long
    On Error Resume Next
    
    ContarReservasEnHabitacionesOcultas = 0
    
    If wsRes Is Nothing Or wsCal Is Nothing Then Exit Function
    
    Dim lastRow As Long, i As Long
    Dim fechaSalida As Date, hoy As Date
    Dim habitaciones As String, arrHabs As Variant
    Dim hab As Variant, filaHab As Long
    Dim contadorOcultas As Long
    
    hoy = Date
    contadorOcultas = 0
    
    lastRow = wsRes.Cells(wsRes.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Function
    
    ' Revisar cada reserva
    For i = 2 To lastRow
        If IsDate(wsRes.Cells(i, "M").Value) And Not IsEmpty(wsRes.Cells(i, "T").Value) Then
            fechaSalida = wsRes.Cells(i, "M").Value
            
            If fechaSalida >= hoy Then
                habitaciones = Trim(CStr(wsRes.Cells(i, "T").Value))
                
                If InStr(habitaciones, ",") > 0 Then
                    arrHabs = Split(habitaciones, ",")
                    Dim resContada As Boolean: resContada = False
                    For Each hab In arrHabs
                        filaHab = MapearHabitacionAFila(Trim(CStr(hab)))
                        If filaHab > 0 Then
                            If wsCal.Rows(filaHab).Hidden Then
                                If Not resContada Then
                                    contadorOcultas = contadorOcultas + 1
                                    resContada = True
                                End If
                                Dim rStr As String
                                rStr = Trim(CStr(hab))
                                If InStr("," & listaOcultas & ",", "," & rStr & ",") = 0 Then
                                    If listaOcultas = "" Then
                                        listaOcultas = rStr
                                    Else
                                        listaOcultas = listaOcultas & ", " & rStr
                                    End If
                                End If
                            End If
                        End If
                    Next hab
                Else
                    filaHab = MapearHabitacionAFila(habitaciones)
                    If filaHab > 0 Then
                        If wsCal.Rows(filaHab).Hidden Then
                            contadorOcultas = contadorOcultas + 1
                            Dim rStr2 As String
                            rStr2 = Trim(CStr(habitaciones))
                            If InStr("," & listaOcultas & ",", "," & rStr2 & ",") = 0 Then
                                If listaOcultas = "" Then
                                    listaOcultas = rStr2
                                Else
                                    listaOcultas = listaOcultas & ", " & rStr2
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    ContarReservasEnHabitacionesOcultas = contadorOcultas
End Function

'=================================================================================
' FUNCI�N: MapearHabitacionAFila
' Convierte c�digo de habitaci�n a n�mero de fila en el calendario
'=================================================================================
Private Function MapearHabitacionAFila(ByVal habitacion As String) As Long
    On Error Resume Next
    MapearHabitacionAFila = 0
    
    Select Case Trim(habitacion)
        Case "1": MapearHabitacionAFila = 2
        Case "2": MapearHabitacionAFila = 4
        Case "3": MapearHabitacionAFila = 6
        Case "4": MapearHabitacionAFila = 8
        Case "5": MapearHabitacionAFila = 10
        Case "6": MapearHabitacionAFila = 12
        Case "7": MapearHabitacionAFila = 14
        Case "8": MapearHabitacionAFila = 16
        Case "9": MapearHabitacionAFila = 18
        Case "10": MapearHabitacionAFila = 20
        Case "11": MapearHabitacionAFila = 22
        Case "12": MapearHabitacionAFila = 24
        Case "13": MapearHabitacionAFila = 26
        Case "14": MapearHabitacionAFila = 28
        Case "15": MapearHabitacionAFila = 30
    End Select
End Function

Private Function EsSolapeGrupoAceptadoEnCelda(ByVal wsCal As Worksheet, ByVal filaHab As Long, ByVal colIndex As Long) As Boolean
    On Error Resume Next
    EsSolapeGrupoAceptadoEnCelda = False
    If wsCal Is Nothing Then Exit Function
    If solapesAceptados Is Nothing Then Exit Function
    If filaHab < 1 Or filaHab > wsCal.Rows.Count Then Exit Function
    If colIndex < 1 Or colIndex > wsCal.Columns.Count Then Exit Function

    Dim dayFirstCol As Long
    dayFirstCol = colIndex - ((colIndex - 2) Mod 3)
    If dayFirstCol < 2 Then dayFirstCol = 2

    Dim fechaCelda As Variant
    fechaCelda = wsCal.Cells(1, dayFirstCol).Value
    If Not IsDate(fechaCelda) Then Exit Function

    Dim habStr As String
    habStr = ObtenerNumHab(filaHab)
    If habStr = "?" Or Trim$(habStr) = "" Then Exit Function

    Dim claveGrupo As String
    claveGrupo = ClaveSolapeGrupoOviedo(habStr)

    EsSolapeGrupoAceptadoEnCelda = SolapeAceptadoOviedo(solapesAceptados, claveGrupo, CDate(fechaCelda))
End Function

