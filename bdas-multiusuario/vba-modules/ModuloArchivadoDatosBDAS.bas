Attribute VB_Name = "ModuloArchivadoDatosBDAS"
' Attribute VB_Name = "ModuloArchivadoDatosBDAS"
'=================================================================================
' M�dulo: Borrado Autom�tico de Datos Personales
' Prop�sito: Cumplimiento normativa registro de viajeros y RGPD
' Fecha: 2025-11-16
' �ltima modificaci�n: 2026-02-12
' Autor: Jaime
' VERSI�N: 3 BDAS separadas (GIJ�N, SOTO, OVIEDO)
'=================================================================================
' BASE LEGAL:
'   Real Decreto 933/2021, de 26 de octubre:
'     "Los sujetos obligados habr�n de llevar un registro inform�tico en el que
'      consten los datos que se relacionan en el Anexo I del Real Decreto 933/2021,
'      de 26 de octubre, incluidos, en su caso, los datos de las personas menores
'      de catorce a�os.
'
'      Los sujetos obligados deber�n registrar y conservar aquellos datos de sus
'      usuarios, que recaben en el ejercicio de su actividad.
'
'      Los datos del registro inform�tico deber�n conservarse durante un plazo de
'      tres a�os a contar desde la finalizaci�n del servicio o prestaci�n contratada."
'
'   RGPD (Reglamento UE 2016/679):
'     Art. 5.1. c - Principio de minimizaci�n de datos
'     Art. 5.1. e - Principio de limitaci�n del plazo de conservaci�n
'
' FUNCIONAMIENTO:
'   - Archiva datos personales de registros > 90 d�as desde FECHA DE SALIDA.
'   - Cada residencia tiene su propia hoja BDAS (GIJ�N, SOTO, OVIEDO).
'   - Los datos se conservan en BDAS durante 3 a�os seg�n RD 933/2021.
'   - Preserva: N� ORDEN, FECHA PETICI�N, N� FACTURA para trazabilidad.
'   - Ejecuta autom�ticamente al abrir el libro (silencioso).
'
' PLAZO DE BORRADO:
'   - HOJAS OPERATIVAS: 90 d�as desde Fecha de Salida (finalizaci�n servicio)
'
' PLAZO DE CONSERVACI�N EN BDAS:
'   - HOJAS BDAS: 3 a�os desde finalizaci�n del servicio (RD 933/2021)
'
' DATOS ARCHIVADOS:
'   - DNI, Nombre, Tel�fono, Direcci�n, Grabaci�n Solicitud
'   - Columnas adicionales: X, Y, Z
'
' DATOS CONSERVADOS EN OPERATIVO:
'   - N� ORDEN, Fecha Petici�n, N� Factura, Fechas estancia, Importes, Estado pago
'=================================================================================

Option Explicit

' Constantes configurables
' Private Const DIAS_RETENCION As Long = 90  ' 90 d�as desde fecha de salida
Private Const NOMBRE_HOJA_CONFIG As String = "CONFIG"
' Nombres de las hojas BDAS
Private Const BDAS_GIJON As String = "BDAS GIJ�N"
Private Const BDAS_SOTO As String = "BDAS SOTO"
Private Const BDAS_OVIEDO As String = "BDAS OVIEDO"
'=================================================================================
' FUNCI�N: Obtener d�as de retenci�n desde CONFIG (configurable por formulario)
' Lee CONFIG!B4. Si no existe o no es v�lido, usa 90 d�as por defecto.
'=================================================================================
Private Function DIAS_RETENCION() As Long
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Dim valor As Variant
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then
        DIAS_RETENCION = 90
        Exit Function
    End If
    
    valor = wsConfig.Range("B4").Value
    
    If IsNumeric(valor) Then
        Dim dias As Long
        dias = CLng(valor)
        ' Respetar los mismos l�mites del formulario (30-365)
        If dias >= 30 And dias <= 365 Then
            DIAS_RETENCION = dias
        Else
            DIAS_RETENCION = 90
        End If
    Else
        DIAS_RETENCION = 90
    End If
    
    On Error GoTo 0
End Function

'=================================================================================
' FUNCION PUBLICA:  Borrado automatico al abrir (SILENCIOSO)
'=================================================================================
Public Sub ArchivarDatosAntiguosAlAbrir()
    On Error Resume Next
    
    Dim totalArchivados As Long
    Dim archiGijon As Long, archiSoto As Long, archiOviedo As Long
    
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    Application.StatusBar = "Iniciando borrado RGPD..."
    
    Call PrepararTodasLasHojasBDAS
    
    Application.StatusBar = "Archivando RESIDENCIA GIJ�N..."
    archiGijon = ArchivarResidencia("RESIDENCIA GIJ�N", BDAS_GIJON, 13, 9, 11, 22, 23, 28, 24, 25, 26, "GIJ�N")
    
    Application.StatusBar = "Archivando RESIDENCIA SOTO..."
    archiSoto = ArchivarResidencia("RESIDENCIA SOTO", BDAS_SOTO, 13, 9, 11, 22, 23, 28, 24, 25, 26, "SOTO")
    
    Application.StatusBar = "Archivando RESIDENCIA OVIEDO..."
    archiOviedo = ArchivarResidencia("RESIDENCIA OVIEDO", BDAS_OVIEDO, 13, 9, 11, 23, 24, 32, 25, 26, 27, "OVIEDO")
    
    totalArchivados = archiGijon + archiSoto + archiOviedo
    
    ' --- Sincronizar RESUMEN tras borrado de filas ---
    If totalArchivados > 0 Then
        Application.StatusBar = "Sincronizando hojas RESUMEN tras archivado..."
        On Error Resume Next
        If archiGijon > 0 Then Call ModuloResumenGijon.ActualizarResumenGijon
        If archiSoto > 0 Then Call ModuloResumenSoto.ActualizarResumenSoto
        If archiOviedo > 0 Then Call ModuloResumenOviedo.ActualizarResumenOviedo
        On Error GoTo 0
    End If
    
    Application.StatusBar = False
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True
    
If totalArchivados > 0 Then
    Call RegistrarArchivadoEnLOG(totalArchivados, archiGijon, archiSoto, archiOviedo)
    
    Dim msg As String
    msg = "Borrado automatico completado." & vbCrLf & vbCrLf
    msg = msg & "Registros procesados:" & vbCrLf
    msg = msg & "  TOTAL: " & totalArchivados & vbCrLf & vbCrLf
    msg = msg & "Datos borrados seg�n normativa vigente." ' & vbCrLf & vbCrLf
    
    MsgBox msg, vbInformation, "Cumplimiento Normativo"
End If
    
    On Error GoTo 0
End Sub

'=================================================================================
' FUNCION PUBLICA:  Borrado manual bajo demanda
' MODIFICADO: Solo archiva la residencia activa (hoja actual)
'=================================================================================
'=================================================================================
' FUNCION PUBLICA: Borrado manual bajo demanda
' VERSION DIAGNOSTICO: Con trazas detalladas
'=================================================================================
Public Sub ArchivarDatosManual()
    On Error GoTo ErrorHandler

    Debug.Print "=== INICIO BorradoDatosManual ==="
    Debug.Print "Timestamp: " & Now

    Dim respuesta As VbMsgBoxResult
    Dim totalArchivados As Long
    Dim hojaActiva As String
    Dim nombreHojaBDAS As String
    Dim origenResidencia As String

    ' >>>>> DIAGN�STICO: Comprobar ActiveSheet
    If ActiveSheet Is Nothing Then
        MsgBox "ERROR: No hay hoja activa.", vbCritical
        Exit Sub
    Else
        Debug.Print "Hoja activa en borrado manual: " & ActiveSheet.Name
    End If

    hojaActiva = ActiveSheet.Name
    Debug.Print "Hoja activa: " & hojaActiva

    ' >>>>> DIAGN�STICO: Verificar existencia de hoja activa en ThisWorkbook
    Dim wsOperativa As Worksheet
    Set wsOperativa = Nothing
    On Error Resume Next
    Set wsOperativa = ThisWorkbook.Sheets(hojaActiva)
    On Error GoTo 0
    If wsOperativa Is Nothing Then
        MsgBox "ERROR: No se encontr� la hoja operativa '" & hojaActiva & "'", vbCritical
        Debug.Print "FALLO: No existe hoja operativa: " & hojaActiva
        Exit Sub
    Else
        Debug.Print "Hoja operativa OK: " & wsOperativa.Name
    End If

    Select Case hojaActiva
        Case "RESIDENCIA GIJ�N"
            nombreHojaBDAS = BDAS_GIJON
            origenResidencia = "GIJ�N"
            Debug.Print "Destino: BDAS GIJ�N"

        Case "RESIDENCIA SOTO"
            nombreHojaBDAS = BDAS_SOTO
            origenResidencia = "SOTO"
            Debug.Print "Destino: BDAS SOTO"

        Case "RESIDENCIA OVIEDO"
            nombreHojaBDAS = BDAS_OVIEDO
            origenResidencia = "OVIEDO"
            Debug.Print "Destino: BDAS OVIEDO"

        Case Else
            Debug.Print "ERROR: Hoja no v�lida"
            MsgBox "El borrado manual solo est� disponible desde las hojas:" & vbCrLf & vbCrLf & _
                   "  - RESIDENCIA GIJ�N" & vbCrLf & _
                   "  - RESIDENCIA SOTO" & vbCrLf & _
                   "  - RESIDENCIA OVIEDO" & vbCrLf & vbCrLf & _
                   "Por favor, sit�ese en una de esas hojas y vuelva a ejecutar el borrado.", _
                   vbExclamation, "Borrado RGPD - Hoja no v�lida"
            Exit Sub
    End Select

    Debug.Print "Preparando mensaje de confirmaci�n..."

    Dim msgConfirm As String
    msgConfirm = "BORRADO DE DATOS PERSONALES" & vbCrLf & vbCrLf
    msgConfirm = msgConfirm & "Se borrar�n los datos de:" & vbCrLf & vbCrLf
    msgConfirm = msgConfirm & "  " & hojaActiva & " " & vbCrLf & vbCrLf
    msgConfirm = msgConfirm & "Registros con m�s de " & DIAS_RETENCION & " d�as desde fecha de salida." & vbCrLf & vbCrLf
    msgConfirm = msgConfirm & "�Continuar?"

    Debug.Print "Mostrando confirmaci�n al usuario..."
    respuesta = MsgBox(msgConfirm, vbQuestion + vbYesNo + vbDefaultButton2, "Borrado de Datos Personales")

    If respuesta = vbNo Then
        Debug.Print "Usuario cancel� el proceso"
        Exit Sub
    End If

    Debug.Print "Usuario confirm� - Iniciando proceso..."

    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual
    Application.Cursor = xlWait

    ' >>>>> DIAGN�STICO: Verificar existencia/creaci�n hoja BDAS
    On Error GoTo ErrorHandler
    Call PrepararHojaBDAS(nombreHojaBDAS)

    ' Comprobar que la hoja BDAS ahora existe
    Dim wsBDAS As Worksheet
    Set wsBDAS = Nothing
    On Error Resume Next
    Set wsBDAS = ThisWorkbook.Sheets(nombreHojaBDAS)
    On Error GoTo 0
    If wsBDAS Is Nothing Then
        MsgBox "ERROR: No se encontr� o no se pudo crear la hoja BDAS '" & nombreHojaBDAS & "'", vbCritical
        Debug.Print "FALLO: No existe hoja BDAS: " & nombreHojaBDAS
        GoTo LimpiezaYFin
    Else
        Debug.Print "Hoja BDAS OK: " & wsBDAS.Name
    End If

    Application.StatusBar = "Procesando " & hojaActiva & "... Por favor espere."

    On Error GoTo ErrorHandler

    Select Case hojaActiva
        Case "RESIDENCIA GIJ�N"
            totalArchivados = ArchivarResidencia("RESIDENCIA GIJ�N", BDAS_GIJON, 13, 9, 11, 22, 23, 28, 24, 25, 26, "GIJ�N")
            
        Case "RESIDENCIA SOTO"
            totalArchivados = ArchivarResidencia("RESIDENCIA SOTO", BDAS_SOTO, 13, 9, 11, 22, 23, 28, 24, 25, 26, "SOTO")
            
        Case "RESIDENCIA OVIEDO"
            totalArchivados = ArchivarResidencia("RESIDENCIA OVIEDO", BDAS_OVIEDO, 13, 9, 11, 23, 24, 32, 25, 26, 27, "OVIEDO")
            
    End Select

LimpiezaYFin:
    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True

    If totalArchivados > 0 Then
        Debug.Print "Registrando en LOG..."
        Call RegistrarArchivadoEnLOG(totalArchivados, _
            IIf(origenResidencia = "GIJ�N", totalArchivados, 0), _
            IIf(origenResidencia = "SOTO", totalArchivados, 0), _
            IIf(origenResidencia = "OVIEDO", totalArchivados, 0))

        ' --- Sincronizar RESUMEN tras borrado manual ---
        Debug.Print "Sincronizando RESUMEN tras borrado manual..."
        On Error Resume Next
        Select Case origenResidencia
            Case "GIJ�N":  Call ModuloResumenGijon.ActualizarResumenGijon
            Case "SOTO":   Call ModuloResumenSoto.ActualizarResumenSoto
            Case "OVIEDO": Call ModuloResumenOviedo.ActualizarResumenOviedo
        End Select
        On Error GoTo 0
        Debug.Print "Sincronizaci�n RESUMEN completada"

        Dim msgResult As String
        msgResult = "Borrado completado con �xito." & vbCrLf & vbCrLf
        msgResult = msgResult & "Registros procesados:" & vbCrLf
        msgResult = msgResult & "  " & hojaActiva & ": " & totalArchivados & " -> " & nombreHojaBDAS & vbCrLf & vbCrLf
        msgResult = msgResult & "Los datos se conservar�n 3 a�os (RD 933/2021)." & vbCrLf & vbCrLf
        msgResult = msgResult & "Cumplimiento normativo OK"

        MsgBox msgResult, vbInformation, "Borrado Completado"
    Else
        Debug.Print "No hay registros para archivar"
        MsgBox "No hay registros para archivar en " & hojaActiva & "." & vbCrLf & vbCrLf & _
               "Todos los registros tienen menos de " & DIAS_RETENCION & " d�as desde la fecha de salida.", _
               vbInformation, "Sin Registros - Cumplimiento OK"
    End If

    Debug.Print "=== FIN ArchivarDatosManual OK ==="
    Exit Sub

ErrorHandler:
    Debug.Print "*** ERROR CAPTURADO ***"
    Debug.Print "N�mero: " & Err.Number
    Debug.Print "Descripci�n: " & Err.Description
    Debug.Print "Fuente: " & Err.Source
    Debug.Print "�ltima l�nea ejecutada visible en traza arriba"

    Application.StatusBar = False
    Application.Cursor = xlDefault
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.screenUpdating = True

    MsgBox "Error durante el borrado:" & vbCrLf & vbCrLf & _
           "N�mero: " & Err.Number & vbCrLf & _
           "Descripci�n: " & Err.Description & vbCrLf & vbCrLf & _
           "Revise la Ventana Inmediato (Ctrl+G) para detalles.", _
           vbCritical, "Error"
End Sub

'=================================================================================
' FUNCION CORE:   Archivar una residencia especifica a su BDAS correspondiente
' ULTRA-OPTIMIZADO: Eliminacion por bloques contiguos (mucho mas rapido que Union)
' OPTIMIZADO: Usa Dictionary para verificacion rapida de duplicados
' MODIFICADO: Elimina filas completas en lugar de solo limpiar celdas
' MODIFICADO: Columna C de BDAS = Columna C de RESIDENCIA, o P si C esta vacia
'=================================================================================
'=================================================================================
' FUNCION CORE: Archivar una residencia espec�fica a su BDAS correspondiente
' CORREGIDO v2: Validaci�n robusta del Dictionary + manejo de errores mejorado
'=================================================================================
Private Function ArchivarResidencia(nombreHojaOrigen As String, _
                                   nombreHojaBDAS As String, _
                                   colFechaSalida As Long, _
                                   colDNI As Long, _
                                   colNombre As Long, _
                                   colTelefono As Long, _
                                   colDireccion As Long, _
                                   colGrabacion As Long, _
                                   colX As Long, _
                                   colY As Long, _
                                   colZ As Long, _
                                   origen As String) As Long
    
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet, wsBDAS As Worksheet
    Dim lastRow As Long, i As Long
    Dim registrosArchivados As Long
    Dim fechaLimite As Date
    Dim nombreLOG As String
    Dim dictBDAS As Object
    
    ' === VALIDACI�N INICIAL ===
    Set ws = Nothing
    Set wsBDAS = Nothing
    
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombreHojaOrigen)
    Set wsBDAS = ThisWorkbook.Worksheets(nombreHojaBDAS)
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Debug.Print "ERROR: Hoja origen '" & nombreHojaOrigen & "' no encontrada"
        ArchivarResidencia = 0
        Exit Function
    End If
    
    If wsBDAS Is Nothing Then
        Debug.Print "ERROR: Hoja BDAS '" & nombreHojaBDAS & "' no encontrada"
        ArchivarResidencia = 0
        Exit Function
    End If
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        ArchivarResidencia = 0
        Exit Function
    End If
    
    fechaLimite = Date - DIAS_RETENCION
    nombreLOG = ObtenerNombreLOGPorOrigen(origen)
    
' === CARGAR DICTIONARY (con fallback) ===
    Set dictBDAS = CargarRegistrosBDAS(wsBDAS)
    
    ' ? MODO ALTERNATIVO: Si Dictionary falla, usar verificaci�n directa
    Dim usarModoDirecto As Boolean
    usarModoDirecto = (dictBDAS Is Nothing)
    
    If usarModoDirecto Then
        Debug.Print "ADVERTENCIA: Usando modo de verificaci�n DIRECTO (m�s lento)"
        Debug.Print "  Recomendaci�n: Habilitar 'Microsoft Scripting Runtime' en Referencias VBA"
    End If
    
    ' === LEER DATOS DE ORIGEN ===
    Dim dataOrigen As Variant
    dataOrigen = ws.Range("A2:AF" & lastRow).Value
    
    Dim dataGrabacion As Variant
    If lastRow > 1 Then
        dataGrabacion = ws.Range(ws.Cells(2, colGrabacion), ws.Cells(lastRow, colGrabacion)).Value
    End If
    
    ' === PREPARAR ARRAYS ===
    Dim datosArchivables() As Variant
    Dim filasAEliminar() As Long
    Dim ordenesEliminados() As Variant
    Dim contadorArchivables As Long
    Dim contadorFilas As Long
    Dim fechaSalida As Variant
    Dim valorColumnaC As Variant
    Dim valorColumnaP As Variant
    Dim valorParaBDAS As Variant
    Dim numOrdenActual As Variant
    Dim fechaPeticionActual As Variant
    Dim esDuplicado As Boolean
    Dim totalFilasOrigen As Long
    
    totalFilasOrigen = UBound(dataOrigen, 1)
    
    ReDim datosArchivables(1 To totalFilasOrigen, 1 To 13)
    ReDim filasAEliminar(1 To totalFilasOrigen)
    ReDim ordenesEliminados(1 To totalFilasOrigen)
    
    contadorArchivables = 0
    contadorFilas = 0
    registrosArchivados = 0
    
    ' === PROCESAR FILAS ===
    For i = 1 To totalFilasOrigen
        fechaSalida = dataOrigen(i, colFechaSalida)
        
        If IsDate(fechaSalida) Then
            If Int(CDate(fechaSalida)) <= Int(fechaLimite) Then
                If Not IsEmpty(dataOrigen(i, colDNI)) Or _
                   Not IsEmpty(dataOrigen(i, colNombre)) Or _
                   Not IsEmpty(dataOrigen(i, colTelefono)) Or _
                   Not IsEmpty(dataOrigen(i, colDireccion)) Then
                    
                    numOrdenActual = dataOrigen(i, 1)
                    fechaPeticionActual = dataOrigen(i, 2)
                    
                    ' ? VERIFICACI�N: Con Dictionary o m�todo directo
                    If usarModoDirecto Then
                        esDuplicado = ExisteRegistroEnBDASDirecto(wsBDAS, numOrdenActual, fechaPeticionActual)
                    Else
                        esDuplicado = ExisteRegistroEnDictionary(dictBDAS, numOrdenActual, fechaPeticionActual)
                    End If
                    
                    contadorFilas = contadorFilas + 1
                    filasAEliminar(contadorFilas) = i + 1
                    ordenesEliminados(contadorFilas) = numOrdenActual
                    
                    If Not esDuplicado Then
                        contadorArchivables = contadorArchivables + 1
                        
                        datosArchivables(contadorArchivables, 1) = dataOrigen(i, 1)
                        datosArchivables(contadorArchivables, 2) = dataOrigen(i, 2)
                        
                        ' Columna C o P
                        valorColumnaC = dataOrigen(i, 3)
                        valorColumnaP = dataOrigen(i, 16)
                        
                        If IsEmpty(valorColumnaC) Or Trim(CStr(valorColumnaC)) = "" Then
                            valorParaBDAS = valorColumnaP
                        Else
                            valorParaBDAS = valorColumnaC
                        End If
                        datosArchivables(contadorArchivables, 3) = valorParaBDAS
                        
                        datosArchivables(contadorArchivables, 4) = dataOrigen(i, colDNI)
                        datosArchivables(contadorArchivables, 5) = dataOrigen(i, colNombre)
                        datosArchivables(contadorArchivables, 6) = dataOrigen(i, colTelefono)
                        datosArchivables(contadorArchivables, 7) = dataOrigen(i, colDireccion)
                        datosArchivables(contadorArchivables, 8) = dataOrigen(i, colX)
                        datosArchivables(contadorArchivables, 9) = dataOrigen(i, colY)
                        datosArchivables(contadorArchivables, 10) = dataOrigen(i, colZ)
                        
                        If IsArray(dataGrabacion) Then
                            datosArchivables(contadorArchivables, 11) = dataGrabacion(i, 1)
                        Else
                            datosArchivables(contadorArchivables, 11) = ""
                        End If
                        
                        datosArchivables(contadorArchivables, 12) = origen
                        datosArchivables(contadorArchivables, 13) = Now
                    End If
                End If
            End If
        End If
    Next i
    
    ' ? LIBERAR DICTIONARY DESPU�S DE USARLO (movido aqu�)
    Set dictBDAS = Nothing
    
    If contadorFilas = 0 Then
        ArchivarResidencia = 0
        Exit Function
    End If
    
    ' === ESCRIBIR EN BDAS ===
    If contadorArchivables > 0 Then
        Dim ultimaFilaBDAS As Long
        ultimaFilaBDAS = wsBDAS.Cells(wsBDAS.Rows.Count, "A").End(xlUp).Row
        If ultimaFilaBDAS = 1 And wsBDAS.Cells(1, 1).Value = "" Then
            Call CrearEncabezadosBDAS(wsBDAS)
            ultimaFilaBDAS = 1
        End If
        
        Dim estabaProtegidaBDAS As Boolean
        estabaProtegidaBDAS = wsBDAS.ProtectContents
        If estabaProtegidaBDAS Then
            On Error Resume Next
            wsBDAS.Unprotect password:=ModuloConfigSegura.ObtenerPasswordHojas()
            On Error GoTo ErrorHandler
        End If
        
        Dim arrayFinal() As Variant
        ReDim arrayFinal(1 To contadorArchivables, 1 To 13)
        For i = 1 To contadorArchivables
            arrayFinal(i, 1) = datosArchivables(i, 1)
            arrayFinal(i, 2) = datosArchivables(i, 2)
            arrayFinal(i, 3) = datosArchivables(i, 3)
            arrayFinal(i, 4) = datosArchivables(i, 4)
            arrayFinal(i, 5) = datosArchivables(i, 5)
            arrayFinal(i, 6) = datosArchivables(i, 6)
            arrayFinal(i, 7) = datosArchivables(i, 7)
            arrayFinal(i, 8) = datosArchivables(i, 8)
            arrayFinal(i, 9) = datosArchivables(i, 9)
            arrayFinal(i, 10) = datosArchivables(i, 10)
            arrayFinal(i, 11) = datosArchivables(i, 11)
            arrayFinal(i, 12) = datosArchivables(i, 12)
            arrayFinal(i, 13) = datosArchivables(i, 13)
        Next i
        
        wsBDAS.Range("A" & (ultimaFilaBDAS + 1)).Resize(contadorArchivables, 13).Value = arrayFinal
        
        If estabaProtegidaBDAS Then
            On Error Resume Next
            wsBDAS.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
            On Error GoTo ErrorHandler
        End If
    End If
    
    ' === ELIMINAR FILAS DE ORIGEN ===
    Dim estabaProtegidaOrigen As Boolean
    estabaProtegidaOrigen = ws.ProtectContents
    
    If estabaProtegidaOrigen Then
        On Error Resume Next
        ws.Unprotect password:=ModuloConfigSegura.ObtenerPasswordHojas()
        On Error GoTo ErrorHandler
    End If
    
    ' Ordenar filas a eliminar (descendente)
    Dim j As Long, temp As Long, tempOrden As Variant
    For i = 1 To contadorFilas - 1
        For j = i + 1 To contadorFilas
            If filasAEliminar(i) < filasAEliminar(j) Then
                temp = filasAEliminar(i)
                filasAEliminar(i) = filasAEliminar(j)
                filasAEliminar(j) = temp
                tempOrden = ordenesEliminados(i)
                ordenesEliminados(i) = ordenesEliminados(j)
                ordenesEliminados(j) = tempOrden
            End If
        Next j
    Next i
    
    ' Registrar en LOG
    If Len(nombreLOG) > 0 Then
        RegistrarEliminacionMasivaEnLOG nombreLOG, ordenesEliminados, contadorFilas, origen
    End If
    
    ' Eliminar filas por bloques
    Dim bloqueInicio As Long
    Dim bloqueFin As Long
    Dim enBloque As Boolean
    
    enBloque = False
    
    For i = 1 To contadorFilas
        If Not enBloque Then
            bloqueInicio = filasAEliminar(i)
            bloqueFin = filasAEliminar(i)
            enBloque = True
        Else
            If filasAEliminar(i) = bloqueFin - 1 Then
                bloqueFin = filasAEliminar(i)
            Else
                ws.Rows(bloqueFin & ":" & bloqueInicio).Delete Shift:=xlUp
                bloqueInicio = filasAEliminar(i)
                bloqueFin = filasAEliminar(i)
            End If
        End If
    Next i
    
    If enBloque Then
        ws.Rows(bloqueFin & ":" & bloqueInicio).Delete Shift:=xlUp
    End If
    
    If estabaProtegidaOrigen Then
        On Error Resume Next
        ws.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
        On Error GoTo ErrorHandler
    End If
    
    registrosArchivados = contadorFilas
    ArchivarResidencia = registrosArchivados
    Exit Function
    
ErrorHandler:
    Debug.Print "ERROR en ArchivarResidencia (" & nombreHojaOrigen & "): " & Err.Number & " - " & Err.Description
    MsgBox "Error durante el borrado de " & nombreHojaOrigen & ":" & vbCrLf & vbCrLf & _
           "N�mero: " & Err.Number & vbCrLf & _
           "Descripci�n: " & Err.Description, vbCritical, "Error de Borrado"
    ArchivarResidencia = 0
End Function

'=================================================================================
' PREPARAR TODAS LAS HOJAS BDAS (3 hojas separadas)
'=================================================================================
Private Sub PrepararTodasLasHojasBDAS()
    On Error Resume Next
    
    ' Preparar BDAS GIJ�N
    Call PrepararHojaBDAS(BDAS_GIJON)
    
    ' Preparar BDAS SOTO
    Call PrepararHojaBDAS(BDAS_SOTO)
    
    ' Preparar BDAS OVIEDO
    Call PrepararHojaBDAS(BDAS_OVIEDO)
    
    On Error GoTo 0
End Sub

'=================================================================================
' PREPARAR UNA HOJA BDAS ESPEC�FICA
' VERSION DIAGNOSTICO
'=================================================================================
Private Sub PrepararHojaBDAS(nombreHoja As String)
    On Error GoTo ErrorHandler
    
    Dim wsBDAS As Worksheet
    Set wsBDAS = Nothing
    
    On Error Resume Next
    Set wsBDAS = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo ErrorHandler
    
    If wsBDAS Is Nothing Then
'        Debug.Print "  Hoja no existe - Creando..."
        Set wsBDAS = ThisWorkbook.Worksheets.Add
        wsBDAS.Name = nombreHoja
'        Debug.Print "  Hoja creada - Creando encabezados..."
        Call CrearEncabezadosBDAS(wsBDAS)
'        Debug.Print "  Encabezados creados"
    Else
'        Debug.Print "  Hoja ya existe"
    End If
    
'    Debug.Print "  Ocultando hoja..."
    wsBDAS.visible = xlSheetVeryHidden
    
'    Debug.Print "  Protegiendo hoja..."
    Dim pwd As String
    pwd = ModuloConfigSegura.ObtenerPasswordHojas()
'    Debug.Print "  Password obtenida: " & IIf(Len(pwd) > 0, "OK (" & Len(pwd) & " chars)", "VACIA!")
    
    wsBDAS.Protect password:=pwd, UserInterfaceOnly:=True
'    Debug.Print "  PrepararHojaBDAS completado OK"
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "  *** ERROR en PrepararHojaBDAS ***"
    Debug.Print "  N�mero: " & Err.Number
    Debug.Print "  Descripci�n: " & Err.Description
    Err.Raise Err.Number, "PrepararHojaBDAS", Err.Description
End Sub
'=================================================================================
' CREAR ENCABEZADOS EN HOJA BDAS
'=================================================================================
Private Sub CrearEncabezadosBDAS(ws As Worksheet)
    On Error Resume Next
    
    With ws
        .Cells(1, 1).Value = "N� ORDEN"
        .Cells(1, 2).Value = "FECHA PETICI�N"
        .Cells(1, 3).Value = "N� FACTURA"
        .Cells(1, 4).Value = "DNI"
        .Cells(1, 5).Value = "NOMBRE"
        .Cells(1, 6).Value = "TEL�FONO"
        .Cells(1, 7).Value = "DIRECCI�N"
        .Cells(1, 8).Value = "COLUMNA X"
        .Cells(1, 9).Value = "COLUMNA Y"
        .Cells(1, 10).Value = "COLUMNA Z"
        .Cells(1, 11).Value = "GRABACI�N"
        .Cells(1, 12).Value = "ORIGEN"
        .Cells(1, 13).Value = "FECHA BORRADO"
        
        .Rows(1).Font.Bold = True
        .Rows(1).Interior.color = RGB(68, 114, 196)
        .Rows(1).Font.color = RGB(255, 255, 255)
        .Columns("A:M").AutoFit
    End With
    
    On Error GoTo 0
End Sub

'=================================================================================
' REGISTRAR BORRADO EN LOG (con desglose por residencia)
'=================================================================================
Private Sub RegistrarArchivadoEnLOG(total As Long, gijon As Long, soto As Long, oviedo As Long)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim mensaje As String
    
    mensaje = "BORRADO HOJAS BDAS: " & total & ""
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJ�N", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            wsLOG.Cells(ultimaFila, 1).Value = "SISTEMA"
            wsLOG.Cells(ultimaFila, 2).Value = Now
            wsLOG.Cells(ultimaFila, 3).Value = mensaje
            wsLOG.Cells(ultimaFila, 4).Value = ""
            wsLOG.Cells(ultimaFila, 5).Value = ""
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
        On Error GoTo 0
    Next i
End Sub

'=================================================================================
' REGISTRAR ELIMINACI�N DE FILA EN LOG (similar a BorradoSolicitudes)
'=================================================================================
Private Sub RegistrarEliminacionFilaEnLOG(ByVal ws As Worksheet, ByVal fila As Long, ByVal nombreLOG As String)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFilaLOG As Long
    Dim numOrden As Variant
    Dim logPassword As String
    
    Set wsLOG = ThisWorkbook.Worksheets(nombreLOG)
    If wsLOG Is Nothing Then Exit Sub
    
    numOrden = ws.Cells(fila, "A").Value
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    
    wsLOG.Unprotect password:=logPassword
    
    ultimaFilaLOG = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
    
    wsLOG.Cells(ultimaFilaLOG, 1).Value = "SISTEMA"
    wsLOG.Cells(ultimaFilaLOG, 2).Value = Now
    wsLOG.Cells(ultimaFilaLOG, 3).Value = "BORRADO FILA"
    wsLOG.Cells(ultimaFilaLOG, 4).Value = "N ORDEN:  " & numOrden
    wsLOG.Cells(ultimaFilaLOG, 5).Value = "Borrado RGPD"
    
    wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
    
    On Error GoTo 0
End Sub

'=================================================================================
' REGISTRAR ELIMINACION MASIVA EN LOG (un solo registro resumido)
'=================================================================================
Private Sub RegistrarEliminacionMasivaEnLOG(ByVal nombreLOG As String, _
                                            ByRef ordenesEliminados() As Variant, _
                                            ByVal totalFilas As Long, _
                                            ByVal origen As String)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFilaLOG As Long
    Dim logPassword As String
    Dim resumenOrdenes As String
    Dim i As Long
    
    Set wsLOG = ThisWorkbook.Worksheets(nombreLOG)
    If wsLOG Is Nothing Then Exit Sub
    
    If totalFilas <= 10 Then
        For i = 1 To totalFilas
            If i > 1 Then resumenOrdenes = resumenOrdenes & ", "
            resumenOrdenes = resumenOrdenes & CStr(ordenesEliminados(i))
        Next i
    Else
        For i = 1 To 5
            If i > 1 Then resumenOrdenes = resumenOrdenes & ", "
            resumenOrdenes = resumenOrdenes & CStr(ordenesEliminados(i))
        Next i
        resumenOrdenes = resumenOrdenes & " ...  "
        For i = totalFilas - 2 To totalFilas
            resumenOrdenes = resumenOrdenes & CStr(ordenesEliminados(i))
            If i < totalFilas Then resumenOrdenes = resumenOrdenes & ", "
        Next i
    End If
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    
    wsLOG.Unprotect password:=logPassword
    
    ultimaFilaLOG = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
    
    wsLOG.Cells(ultimaFilaLOG, 1).Value = "SISTEMA"
    wsLOG.Cells(ultimaFilaLOG, 2).Value = Now
    wsLOG.Cells(ultimaFilaLOG, 3).Value = "BORRADO MASIVO"
    wsLOG.Cells(ultimaFilaLOG, 4).Value = totalFilas & " filas eliminadas - Ordenes: " & resumenOrdenes
    wsLOG.Cells(ultimaFilaLOG, 5).Value = "Borrado RGPD - " & origen
    
    wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
    
    On Error GoTo 0
End Sub

'=================================================================================
' OBTENER NOMBRE DE HOJA LOG SEG�N ORIGEN
'=================================================================================
Private Function ObtenerNombreLOGPorOrigen(ByVal origen As String) As String
    Select Case UCase(origen)
        Case "GIJON"
            ObtenerNombreLOGPorOrigen = "LOG_GIJ�N"
        Case "SOTO"
            ObtenerNombreLOGPorOrigen = "LOG_SOTO"
        Case "OVIEDO"
            ObtenerNombreLOGPorOrigen = "LOG_OVIEDO"
        Case Else
            ObtenerNombreLOGPorOrigen = ""
    End Select
End Function

'=================================================================================
' CARGAR REGISTROS EXISTENTES EN BDAS A UN DICTIONARY
' CORREGIDO v3: Manejo robusto de errores + fallback sin Dictionary
'=================================================================================
Private Function CargarRegistrosBDAS(ByVal wsBDAS As Worksheet) As Object
    Dim dict As Object
    
    ' ? INTENTAR CREAR DICTIONARY CON MANEJO EXPL�CITO
    On Error GoTo ErrorCrearDictionary
    Set dict = CreateObject("Scripting.Dictionary")
    On Error GoTo 0
    
    ' Validar creaci�n exitosa
    If dict Is Nothing Then
        Debug.Print "ADVERTENCIA: No se pudo crear Dictionary - usando verificaci�n alternativa"
        Set CargarRegistrosBDAS = Nothing
        Exit Function
    End If
    
    ' Validar hoja BDAS
    If wsBDAS Is Nothing Then
        Set CargarRegistrosBDAS = dict
        Exit Function
    End If
    
    On Error GoTo ErrorProcesar
    
    Dim ultimaFila As Long
    ultimaFila = wsBDAS.Cells(wsBDAS.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFila < 2 Then
        Set CargarRegistrosBDAS = dict
        Exit Function
    End If
    
    Dim datosBDAS As Variant
    datosBDAS = wsBDAS.Range("A2:B" & ultimaFila).Value
    
    Dim i As Long
    Dim clave As String
    Dim numOrden As Variant
    Dim fechaPeticion As Variant
    
    For i = 1 To UBound(datosBDAS, 1)
        numOrden = datosBDAS(i, 1)
        fechaPeticion = datosBDAS(i, 2)
        
        If Not IsEmpty(numOrden) Then
            If IsDate(fechaPeticion) Then
                clave = CStr(numOrden) & "|" & Format(CDate(fechaPeticion), "YYYYMMDD")
            Else
                clave = CStr(numOrden) & "|" & CStr(fechaPeticion)
            End If
            
            If Not dict.Exists(clave) Then
                dict.Add clave, True
            End If
        End If
    Next i
    
    Set CargarRegistrosBDAS = dict
    Exit Function
    
ErrorCrearDictionary:
    Debug.Print "ERROR al crear Dictionary: " & Err.Number & " - " & Err.Description
    Debug.Print "Posible causa: Referencia 'Microsoft Scripting Runtime' no disponible"
    Set CargarRegistrosBDAS = Nothing
    Exit Function
    
ErrorProcesar:
    Debug.Print "ERROR al procesar BDAS: " & Err.Description
    ' Devolver dictionary parcial en caso de error
    Set CargarRegistrosBDAS = dict
End Function

'=================================================================================
' VERIFICAR SI UN REGISTRO YA EXISTE EN EL DICTIONARY DE BDAS
' Compara N� ORDEN (columna A) y FECHA PETICION (columna B)
' Devuelve True si el registro ya existe, False si no existe
'=================================================================================
Private Function ExisteRegistroEnDictionary(ByVal dict As Object, _
                                            ByVal numOrden As Variant, _
                                            ByVal fechaPeticion As Variant) As Boolean
    On Error Resume Next
    
    Dim clave As String
    
    ExisteRegistroEnDictionary = False
    
    If dict Is Nothing Then Exit Function
    If IsEmpty(numOrden) Then Exit Function
    
    If IsDate(fechaPeticion) Then
        clave = CStr(numOrden) & "|" & Format(CDate(fechaPeticion), "YYYYMMDD")
    Else
        clave = CStr(numOrden) & "|" & CStr(fechaPeticion)
    End If
    
    ExisteRegistroEnDictionary = dict.Exists(clave)
    
    On Error GoTo 0
End Function


'=================================================================================
' VERIFICAR DUPLICADO SIN DICTIONARY (m�todo alternativo m�s lento)
' Se usa solo si Dictionary no est� disponible
'=================================================================================
Private Function ExisteRegistroEnBDASDirecto(ByVal wsBDAS As Worksheet, _
                                             ByVal numOrden As Variant, _
                                             ByVal fechaPeticion As Variant) As Boolean
    On Error Resume Next
    
    ExisteRegistroEnBDASDirecto = False
    
    If wsBDAS Is Nothing Then Exit Function
    If IsEmpty(numOrden) Then Exit Function
    
    Dim ultimaFila As Long
    ultimaFila = wsBDAS.Cells(wsBDAS.Rows.Count, "A").End(xlUp).Row
    
    If ultimaFila < 2 Then Exit Function
    
    Dim i As Long
    Dim ordenBDAS As Variant
    Dim fechaBDAS As Variant
    Dim claveBuscada As String
    Dim claveBDAS As String
    
    ' Preparar clave buscada
    If IsDate(fechaPeticion) Then
        claveBuscada = CStr(numOrden) & "|" & Format(CDate(fechaPeticion), "YYYYMMDD")
    Else
        claveBuscada = CStr(numOrden) & "|" & CStr(fechaPeticion)
    End If
    
    ' Buscar en BDAS (m�s lento que Dictionary pero funciona siempre)
    For i = 2 To ultimaFila
        ordenBDAS = wsBDAS.Cells(i, 1).Value
        fechaBDAS = wsBDAS.Cells(i, 2).Value
        
        If Not IsEmpty(ordenBDAS) Then
            If IsDate(fechaBDAS) Then
                claveBDAS = CStr(ordenBDAS) & "|" & Format(CDate(fechaBDAS), "YYYYMMDD")
            Else
                claveBDAS = CStr(ordenBDAS) & "|" & CStr(fechaBDAS)
            End If
            
            If claveBDAS = claveBuscada Then
                ExisteRegistroEnBDASDirecto = True
                Exit Function
            End If
        End If
    Next i
    
    On Error GoTo 0
End Function

'=================================================================================
' BUSCAR EN BDAS (busca en las 3 hojas BDAS)
'=================================================================================
Public Function buscarEnBDAS(criterioBusqueda As String, tipoBusqueda As String) As Collection
    Dim resultados As Collection
    Set resultados = New Collection
    
    Dim hojasBDAS As Variant
    Dim i As Long
    
    hojasBDAS = Array(BDAS_GIJON, BDAS_SOTO, BDAS_OVIEDO)
    
    ' Buscar en cada BDAS
    For i = LBound(hojasBDAS) To UBound(hojasBDAS)
        Dim resultadosHoja As Collection
        Set resultadosHoja = BuscarEnHojaBDAS(CStr(hojasBDAS(i)), criterioBusqueda, tipoBusqueda)
        
        ' Agregar resultados a la colecci�n principal
        Dim item As Variant
        For Each item In resultadosHoja
            resultados.Add item
        Next item
    Next i
    
    Set buscarEnBDAS = resultados
End Function

'=================================================================================
' BUSCAR EN UNA HOJA BDAS ESPEC�FICA
'=================================================================================
Private Function BuscarEnHojaBDAS(nombreHoja As String, criterioBusqueda As String, tipoBusqueda As String) As Collection
    Dim wsBDAS As Worksheet
    Dim resultados As Collection
    Dim lastRow As Long, i As Long
    Dim valorCelda As String
    Dim colBusqueda As Long
    Dim resultado(1 To 13) As Variant
    
    Set resultados = New Collection
    
    On Error Resume Next
    Set wsBDAS = ThisWorkbook.Worksheets(nombreHoja)
    On Error GoTo 0
    
    If wsBDAS Is Nothing Then
        Set BuscarEnHojaBDAS = resultados
        Exit Function
    End If
    
    Select Case UCase(tipoBusqueda)
        Case "ORDEN": colBusqueda = 1
        Case "FACTURA": colBusqueda = 3
        Case "DNI": colBusqueda = 4
        Case "NOMBRE": colBusqueda = 5
        Case Else: colBusqueda = 1
    End Select
    
    lastRow = wsBDAS.Cells(wsBDAS.Rows.Count, colBusqueda).End(xlUp).Row
    
    If lastRow < 2 Then
        Set BuscarEnHojaBDAS = resultados
        Exit Function
    End If
    
    For i = 2 To lastRow
        valorCelda = UCase(Trim(CStr(wsBDAS.Cells(i, colBusqueda).Value)))
        
        If InStr(valorCelda, UCase(Trim(criterioBusqueda))) > 0 Then
            resultado(1) = wsBDAS.Cells(i, 1).Value
            resultado(2) = wsBDAS.Cells(i, 2).Value
            resultado(3) = wsBDAS.Cells(i, 3).Value
            resultado(4) = wsBDAS.Cells(i, 4).Value
            resultado(5) = wsBDAS.Cells(i, 5).Value
            resultado(6) = wsBDAS.Cells(i, 6).Value
            resultado(7) = wsBDAS.Cells(i, 7).Value
            resultado(8) = wsBDAS.Cells(i, 8).Value
            resultado(9) = wsBDAS.Cells(i, 9).Value
            resultado(10) = wsBDAS.Cells(i, 10).Value
            resultado(11) = wsBDAS.Cells(i, 11).Value
            resultado(12) = wsBDAS.Cells(i, 12).Value
            resultado(13) = wsBDAS.Cells(i, 13).Value
            
            resultados.Add resultado
        End If
    Next i
    
    Set BuscarEnHojaBDAS = resultados
End Function

'=================================================================================
' OBTENER NOMBRES DE HOJAS BDAS (funci�n auxiliar)
'=================================================================================
Public Function ObtenerNombresHojasBDAS() As Variant
    ObtenerNombresHojasBDAS = Array(BDAS_GIJON, BDAS_SOTO, BDAS_OVIEDO)
End Function

'=================================================================================
' OBTENER CONFIGURACI�N
'=================================================================================
Public Function BuscarBDASActivado() As Boolean
    On Error Resume Next
    
    Dim wsConfig As Worksheet
    Dim valorConfig As Variant
    
    Set wsConfig = ThisWorkbook.Worksheets(NOMBRE_HOJA_CONFIG)
    
    If wsConfig Is Nothing Then
        BuscarBDASActivado = True
        Exit Function
    End If
    
    valorConfig = wsConfig.Range("B5").Value
    
    If IsEmpty(valorConfig) Then
        BuscarBDASActivado = True
    ElseIf UCase(Trim(CStr(valorConfig))) = "SI" Or UCase(Trim(CStr(valorConfig))) = "S�" Then
        BuscarBDASActivado = True
    Else
        BuscarBDASActivado = False
    End If
    
    On Error GoTo 0
End Function

'=================================================================================
' APLICAR UserInterfaceOnly
'=================================================================================
Public Sub AplicarUserInterfaceOnlyATodasLasHojas()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim hojas As Variant
    Dim nombreHoja As Variant
    
    hojas = Array("RESIDENCIA GIJ�N", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO")
    
    For Each nombreHoja In hojas
        Set ws = ThisWorkbook.Worksheets(CStr(nombreHoja))
        
        If Not ws Is Nothing Then
            If ws.ProtectContents Then
                ws.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), _
                           UserInterfaceOnly:=True, _
                           AllowFormattingCells:=True, _
                           AllowSorting:=True, _
                           AllowFiltering:=True
            End If
        End If
    Next nombreHoja
    
    On Error GoTo 0
End Sub

'=================================================================================
' CONTAR REGISTROS PARA ARCHIVAR
'=================================================================================
Public Function ContarRegistrosParaArchivar() As Long
    On Error Resume Next
    
    Dim total As Long
    total = 0
    
    total = total + ContarRegistrosResidencia("RESIDENCIA GIJ�N", 13, 9, 11, 22, 23)
    total = total + ContarRegistrosResidencia("RESIDENCIA SOTO", 13, 9, 11, 22, 23)
    total = total + ContarRegistrosResidencia("RESIDENCIA OVIEDO", 13, 9, 11, 23, 24)
    
    ContarRegistrosParaArchivar = total
End Function

'=================================================================================
' CONTAR REGISTROS DE UNA RESIDENCIA
'=================================================================================
Private Function ContarRegistrosResidencia(nombreHoja As String, _
                                          colFechaSalida As Long, _
                                          colDNI As Long, _
                                          colNombre As Long, _
                                          colTelefono As Long, _
                                          colDireccion As Long) As Long
    On Error Resume Next
    
    Dim ws As Worksheet
    Set ws = Nothing
    Set ws = ThisWorkbook.Sheets(nombreHoja)
    
'---BORRAR, CODIGO PARA LOCALIZAR ERROR ----- BORRAR, CODIGO PARA LOCALIZAR ERROR -------
If ws Is Nothing Then
    MsgBox "ERROR: No se encontr� la hoja '" & nombreHoja & "'", vbCritical
Else
    Debug.Print "Objeto ws (hoja operativa) cargado correctamente: " & ws.Name
End If
'---BORRAR, CODIGO PARA LOCALIZAR ERROR ----- BORRAR, CODIGO PARA LOCALIZAR ERROR -------
    
    If ws Is Nothing Then
        ContarRegistrosResidencia = 0
        Exit Function
    End If
    
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    If ultimaFila <= 1 Then
        ContarRegistrosResidencia = 0
        Exit Function
    End If
    
    Dim fechaCorte As Date
    fechaCorte = Date - DIAS_RETENCION
    
    Dim i As Long
    Dim contador As Long
    contador = 0
    
    For i = 2 To ultimaFila
        Dim fechaSalida As Variant
        fechaSalida = ws.Cells(i, colFechaSalida).Value
        
        If IsDate(fechaSalida) Then
            If Int(CDate(fechaSalida)) <= Int(fechaCorte) Then
                If Not IsEmpty(ws.Cells(i, colDNI).Value) Or _
                   Not IsEmpty(ws.Cells(i, colNombre).Value) Or _
                   Not IsEmpty(ws.Cells(i, colTelefono).Value) Or _
                   Not IsEmpty(ws.Cells(i, colDireccion).Value) Then
                    contador = contador + 1
                End If
            End If
        End If
    Next i
    
    ContarRegistrosResidencia = contador
End Function
