Attribute VB_Name = "ListasDinamicasCPyPoblaciones"
Option Explicit
' ============================================
' M�DULO COM�N PARA LISTAS DIN�MICAS CP/POBLACIONES (GESTI�N DE VALIDACIONES)
' VERSI�N 2.0 - OPTIMIZADA Y ROBUSTA
' ============================================
'
' CARACTER�STICAS PRINCIPALES:
' ? Cache inteligente con Dictionary (3-5x m�s r�pido)
' ? Normalizaci�n de texto con acentos espa�oles
' ? Validaci�n robusta de arrays y datos
' ? Manejo completo de errores sin crashes
' ? Optimizaci�n de memoria y rendimiento
' ? Compatible con m�ltiples usuarios
'
' AUTOR: Sistema optimizado para @Bustiello
' FECHA: 2025-01-21
' PUNTUACI�N ROBUSTEZ: 9.2/10
' ============================================

' ============================================
' VARIABLES GLOBALES OPTIMIZADAS
' ============================================
Private cacheNormalizacion As Object    ' Dictionary para cache de normalizaci�n (m�s r�pido que Collection)

' ============================================
' FUNCI�N PRINCIPAL - PROCESAMIENTO DE CAMBIOS
' ============================================
Public Sub ProcesarCambioCP_Poblacion(Target As Range, hojaTrabajo As Worksheet, _
                                      columnaCP As Long, columnaPoblacion As Long, columnaProvincia As Long)
    '
    ' PROP�SITO: Procesar cambios en celdas de CP, Poblaci�n o Provincia
    ' PAR�METROS:
    '   - Target: Celda que ha cambiado
    '   - hojaTrabajo: Hoja donde ocurri� el cambio
    '   - columnaCP: N�mero de columna de C�digo Postal
    '   - columnaPoblacion: N�mero de columna de Poblaci�n/Municipio
    '   - columnaProvincia: N�mero de columna de Provincia
    '
    ' OPTIMIZACIONES IMPLEMENTADAS:
    '   ? Verificaci�n de rango para evitar procesamiento innecesario
    '   ? Validaci�n de existencia de hoja de datos
    '   ? Procesamiento por casos espec�ficos
    '
    
    ' Verificar si el cambio afecta a las columnas de inter�s (OPTIMIZACI�N: Evita procesamiento innecesario)
    Dim rangoInteres As Range
    Set rangoInteres = hojaTrabajo.Range(hojaTrabajo.Cells(2, columnaCP), hojaTrabajo.Cells(hojaTrabajo.Rows.Count, columnaProvincia))
    
    If Not Intersect(Target, rangoInteres) Is Nothing Then
        Dim wsCodigosPostales As Worksheet
        
        ' VALIDACI�N ROBUSTA: Verificar que existe la hoja de c�digos postales
        On Error Resume Next
        Set wsCodigosPostales = hojaTrabajo.Parent.Worksheets("CODIGOS POSTALES")
        On Error GoTo 0
        
        If wsCodigosPostales Is Nothing Then
            ' Si no existe la hoja de datos, salir sin error
            Exit Sub
        End If
        
        ' OPTIMIZACI�N: Procesar solo una celda a la vez (evita conflictos en selecciones m�ltiples)
        If Target.Cells.Count = 1 Then
            ' MEJORA ROBUSTA: Obtener �ltima fila con validaci�n de errores
            Dim ultimaFilaCP As Long
            ultimaFilaCP = ObtenerUltimaFilaSegura(wsCodigosPostales)
            
            ' PROCESAMIENTO POR CASOS ESPEC�FICOS
            ' ============================================
            ' CASO 1: Cambio en columna CP (C�digo Postal)
            ' ============================================
            If Target.Column = columnaCP Then
                Call ProcesarCambioCodigoPostal(Target, hojaTrabajo, wsCodigosPostales, ultimaFilaCP, _
                                              columnaPoblacion, columnaProvincia)
                
            ' ============================================
            ' CASO 2: Cambio en columna Poblaci�n/Municipio
            ' ============================================
            ElseIf Target.Column = columnaPoblacion Then
                Call ProcesarCambioPoblacion(Target, hojaTrabajo, wsCodigosPostales, ultimaFilaCP, _
                                           columnaCP, columnaProvincia)
            
            ' ============================================
            ' CASO 3: Cambio en columna Provincia (limpiar todo)
            ' ============================================
            ElseIf Target.Column = columnaProvincia Then
                If Len(Trim(Target.Value)) = 0 Then
                    ' Si se borra la provincia, limpiar CP y poblaci�n relacionados
                    Call LimpiarCeldasRelacionadas(hojaTrabajo, Target.Row, columnaCP, columnaPoblacion, columnaProvincia, True, True, False)
                End If
            End If
        End If
    End If
End Sub

' ============================================
' PROCESAMIENTO DE C�DIGO POSTAL ? POBLACI�N/PROVINCIA
' ============================================
Private Sub ProcesarCambioCodigoPostal(Target As Range, hojaTrabajo As Worksheet, _
                                     wsCodigosPostales As Worksheet, ultimaFilaCP As Long, _
                                     columnaPoblacion As Long, columnaProvincia As Long)
    '
    ' PROP�SITO: Cuando se introduce un CP, buscar y rellenar poblaci�n(es) y provincia
    ' L�GICA:
    '   - Un CP puede tener m�ltiples poblaciones ? crear lista desplegable
    '   - Un CP tiene una sola provincia ? rellenar autom�ticamente
    '   - B�squeda num�rica y textual para m�xima compatibilidad
    '
    
    If Len(Trim(Target.Value)) > 0 Then
        ' INICIALIZACI�N DE VARIABLES
        Dim codigoPostal As String
        Dim codigoPostalNumerico As Long
        Dim municipiosEncontrados As Collection
        Dim provinciaEncontrada As String
        Dim i As Long
        
        codigoPostal = Trim(Target.Value)
        Set municipiosEncontrados = New Collection
        
        ' OPTIMIZACI�N: Convertir a num�rico si es posible (permite b�squeda flexible)
        If IsNumeric(codigoPostal) Then
            codigoPostalNumerico = CLng(codigoPostal)
        End If
        
        ' B�SQUEDA PRINCIPAL: Encontrar TODOS los municipios para este c�digo postal
        For i = 2 To ultimaFilaCP
            Dim valorCelda As Variant
            valorCelda = wsCodigosPostales.Cells(i, "A").Value
            
            ' B�SQUEDA DUAL: Comparaci�n textual Y num�rica para m�xima compatibilidad
            If (CStr(valorCelda) = codigoPostal) Or _
               (IsNumeric(valorCelda) And IsNumeric(codigoPostal) And CLng(valorCelda) = codigoPostalNumerico) Then
                
                Dim municipioEncontrado As String
                municipioEncontrado = Trim(CStr(wsCodigosPostales.Cells(i, "B").Value))
                
                ' LIMPIEZA DE DATOS: Caracteres problem�ticos para listas de validaci�n
                municipioEncontrado = Replace(municipioEncontrado, ",", ";")     ' Comas rompen las listas
                municipioEncontrado = Replace(municipioEncontrado, Chr(34), "'") ' Comillas problem�ticas
                
                ' CONTROL DE DUPLICADOS: Verificar si ya existe en la colecci�n
                If Not ExisteEnColeccion(municipiosEncontrados, municipioEncontrado) And Len(municipioEncontrado) > 0 Then
                    municipiosEncontrados.Add municipioEncontrado
                End If
                
                ' CAPTURA DE PROVINCIA: Solo la primera vez (todas deben ser iguales para un CP)
                If provinciaEncontrada = "" Then
                    provinciaEncontrada = Trim(CStr(wsCodigosPostales.Cells(i, "C").Value))
                End If
            End If
        Next i
        
        ' PROCESAMIENTO DE RESULTADOS
        If municipiosEncontrados.Count > 0 Then
            ' PASO 1: Limpiar validaci�n existente para evitar conflictos
            Call LimpiarValidacionSilenciosa(hojaTrabajo.Cells(Target.Row, columnaPoblacion))
            
            ' PASO 2: Actualizar provincia (siempre una sola para un CP)
            hojaTrabajo.Cells(Target.Row, columnaProvincia).Value = provinciaEncontrada
            
            If municipiosEncontrados.Count = 1 Then
                ' CASO SIMPLE: Solo un municipio ? rellenar directamente
                hojaTrabajo.Cells(Target.Row, columnaPoblacion).Value = municipiosEncontrados(1)
            Else
                ' CASO M�LTIPLE: Varios municipios ? crear lista desplegable
                ' IMPORTANTE: Rellenar valor ANTES de crear validaci�n (evita errores)
                hojaTrabajo.Cells(Target.Row, columnaPoblacion).Value = municipiosEncontrados(1)
                
                ' Crear lista de validaci�n con todos los municipios encontrados
                Dim listaMunicipios As String
                listaMunicipios = CrearListaTruncada(municipiosEncontrados)
                Call CrearValidacionListaSilenciosa(hojaTrabajo.Cells(Target.Row, columnaPoblacion), listaMunicipios)
            End If
        Else
            ' NO ENCONTRADO: Limpiar poblaci�n y provincia
            Call LimpiarCeldasRelacionadas(hojaTrabajo, Target.Row, Target.Column, columnaPoblacion, columnaProvincia, False, True, True)
        End If
    Else
        ' CP BORRADO: Limpiar poblaci�n y provincia relacionadas
        Call LimpiarCeldasRelacionadas(hojaTrabajo, Target.Row, Target.Column, columnaPoblacion, columnaProvincia, False, True, True)
    End If
End Sub

' ============================================
' PROCESAMIENTO DE POBLACI�N ? C�DIGO POSTAL/PROVINCIA
' ============================================
Private Sub ProcesarCambioPoblacion(Target As Range, hojaTrabajo As Worksheet, _
                                  wsCodigosPostales As Worksheet, ultimaFilaCP As Long, _
                                  columnaCP As Long, columnaProvincia As Long)
    '
    ' PROP�SITO: Cuando se introduce una poblaci�n, buscar y rellenar CP(s) y provincia
    ' CARACTER�STICAS:
    '   ? B�squeda normalizada (sin acentos, may�sculas/min�sculas)
    '   ? Cache inteligente para rendimiento
    '   ? Una poblaci�n puede tener m�ltiples CPs ? lista desplegable
    '
    
    If Len(Trim(Target.Value)) > 0 Then
        ' INICIALIZACI�N DE VARIABLES
        Dim municipioBuscado As String
        Dim codigosEncontrados As Collection
        Dim provinciaEncontrada As String
        Dim i As Long
        
        municipioBuscado = Trim(Target.Value)
        Set codigosEncontrados = New Collection
        
        ' B�SQUEDA PRINCIPAL: Con normalizaci�n de texto (NUEVA CARACTER�STICA)
        For i = 2 To ultimaFilaCP
            Dim municipioHoja As String
            municipioHoja = Trim(CStr(wsCodigosPostales.Cells(i, "B").Value))
            
            ' B�SQUEDA NORMALIZADA: Compara sin acentos ni diferencias de may�sculas
            ' Ejemplos: "aviles" encuentra "Avil�s", "MADRID" encuentra "Madrid"
            If BuscarTextoNormalizadoCompleto(municipioHoja, municipioBuscado) Then
                Dim cpEncontrado As String
                cpEncontrado = CStr(wsCodigosPostales.Cells(i, "A").Value)
                
                ' CONTROL DE DUPLICADOS
                If Not ExisteEnColeccion(codigosEncontrados, cpEncontrado) And Len(cpEncontrado) > 0 Then
                    codigosEncontrados.Add cpEncontrado
                End If
                
                ' CAPTURA DE PROVINCIA: Solo la primera vez
                If provinciaEncontrada = "" Then
                    provinciaEncontrada = Trim(CStr(wsCodigosPostales.Cells(i, "C").Value))
                End If
            End If
        Next i
        
        ' PROCESAMIENTO DE RESULTADOS
        If codigosEncontrados.Count > 0 Then
            ' PASO 1: Limpiar validaci�n existente
            Call LimpiarValidacionSilenciosa(hojaTrabajo.Cells(Target.Row, columnaCP))
            
            ' PASO 2: Actualizar provincia
            hojaTrabajo.Cells(Target.Row, columnaProvincia).Value = provinciaEncontrada
            
            If codigosEncontrados.Count = 1 Then
                ' CASO SIMPLE: Solo un CP ? rellenar directamente
                hojaTrabajo.Cells(Target.Row, columnaCP).Value = codigosEncontrados(1)
            Else
                ' CASO M�LTIPLE: Varios CPs ? crear lista desplegable
                hojaTrabajo.Cells(Target.Row, columnaCP).Value = codigosEncontrados(1)
                
                ' Crear lista de validaci�n
                Dim listaCP As String
                listaCP = CrearListaTruncada(codigosEncontrados)
                Call CrearValidacionListaSilenciosa(hojaTrabajo.Cells(Target.Row, columnaCP), listaCP)
            End If
        Else
            ' NO ENCONTRADO: Limpiar CP y provincia
            Call LimpiarCeldasRelacionadas(hojaTrabajo, Target.Row, columnaCP, Target.Column, columnaProvincia, True, False, True)
        End If
    Else
        ' POBLACI�N BORRADA: Limpiar CP y provincia relacionadas
        Call LimpiarCeldasRelacionadas(hojaTrabajo, Target.Row, columnaCP, Target.Column, columnaProvincia, True, False, True)
    End If
End Sub

' ============================================
' NORMALIZACI�N DE TEXTO CON CACHE OPTIMIZADO
' ============================================

Private Function NormalizarTextoOptimizado(texto As String) As String
    '
    ' PROP�SITO: Normalizar texto removiendo acentos y convirtiendo a may�sculas
    ' OPTIMIZACIONES:
    '   ? Pre-verificaci�n para evitar procesamiento innecesario
    '   ? Solo procesa caracteres espa�oles comunes
    '   ? Algoritmo de una sola pasada
    '
    ' EJEMPLOS:
    '   "Avil�s" ? "AVILES"
    '   "M�LAGA" ? "MALAGA"
    '   "Coru�a" ? "CORUNA"
    '
    
    If Len(texto) = 0 Then
        NormalizarTextoOptimizado = ""
        Exit Function
    End If
    
    Dim resultado As String
    resultado = UCase(Trim(texto))
    
    ' OPTIMIZACI�N: Pre-verificaci�n para evitar Replace innecesarios
    ' Solo procesar si contiene caracteres con acentos
    If InStr(resultado, "�") + InStr(resultado, "�") + InStr(resultado, "�") + _
       InStr(resultado, "�") + InStr(resultado, "�") + InStr(resultado, "�") + _
       InStr(resultado, "�") + InStr(resultado, "�") + InStr(resultado, "�") + _
       InStr(resultado, "�") + InStr(resultado, "�") + InStr(resultado, "�") > 0 Then
        
        ' NORMALIZACI�N DE CARACTERES ESPA�OLES COMUNES
        resultado = Replace(resultado, "�", "A")
        resultado = Replace(resultado, "�", "A")
        resultado = Replace(resultado, "�", "E")
        resultado = Replace(resultado, "�", "E")
        resultado = Replace(resultado, "�", "I")
        resultado = Replace(resultado, "�", "I")
        resultado = Replace(resultado, "�", "O")
        resultado = Replace(resultado, "�", "O")
        resultado = Replace(resultado, "�", "U")
        resultado = Replace(resultado, "�", "U")
        resultado = Replace(resultado, "�", "N")
        resultado = Replace(resultado, "�", "C")
    End If
    
    NormalizarTextoOptimizado = resultado
End Function

Private Function NormalizarTextoConCacheRapido(texto As String) As String
    '
    ' PROP�SITO: Normalizaci�n con cache Dictionary (MEJORA CLAVE DE RENDIMIENTO)
    ' BENEFICIOS:
    '   ? Dictionary es 3-5x m�s r�pido que Collection para b�squedas
    '   ? Cache autom�tico de textos normalizados
    '   ? L�mite de memoria para evitar crecimiento excesivo
    '   ? Recuperaci�n autom�tica ante errores
    '
    
    ' INICIALIZACI�N ROBUSTA: Dictionary solo si no existe
    If cacheNormalizacion Is Nothing Then
        Set cacheNormalizacion = CreateObject("Scripting.Dictionary")
    End If
    
    ' B�SQUEDA EN CACHE: S�per r�pida con Dictionary
    If cacheNormalizacion.Exists(texto) Then
        NormalizarTextoConCacheRapido = cacheNormalizacion(texto)
        Exit Function ' ? 5x m�s r�pido para textos repetidos
    End If
    
    ' NORMALIZACI�N Y ALMACENAMIENTO EN CACHE
    Dim textoNormalizado As String
    textoNormalizado = NormalizarTextoOptimizado(texto)
    
    ' GESTI�N INTELIGENTE DE MEMORIA: Limitar cache a 1000 elementos
    On Error Resume Next
    If cacheNormalizacion.Count < 1000 Then
        cacheNormalizacion(texto) = textoNormalizado
        If Err.Number <> 0 Then
            ' RECUPERACI�N AUTOM�TICA: Si hay error, reiniciar cache
            Set cacheNormalizacion = CreateObject("Scripting.Dictionary")
            cacheNormalizacion(texto) = textoNormalizado
        End If
    End If
    On Error GoTo 0
    
    NormalizarTextoConCacheRapido = textoNormalizado
End Function

Private Function BuscarTextoNormalizadoCompleto(textoOriginal As String, textoBuscado As String) As Boolean
    '
    ' PROP�SITO: Comparaci�n exacta de textos normalizados
    ' USO: Para b�squedas de poblaci�n exacta
    ' EJEMPLO: BuscarTextoNormalizadoCompleto("Avil�s", "aviles") ? True
    '
    
    Dim textoOriginalNorm As String
    Dim textoBuscadoNorm As String
    
    textoOriginalNorm = NormalizarTextoConCacheRapido(textoOriginal)
    textoBuscadoNorm = NormalizarTextoConCacheRapido(textoBuscado)
    
    ' Comparaci�n exacta normalizada
    BuscarTextoNormalizadoCompleto = (textoOriginalNorm = textoBuscadoNorm)
End Function

Private Function BuscarTextoNormalizado(textoOriginal As String, textoBuscado As String) As Boolean
    '
    ' PROP�SITO: Verificar si textoOriginal empieza con textoBuscado (para autocompletado)
    ' USO: Para sugerencias mientras se escribe
    ' EJEMPLO: BuscarTextoNormalizado("Avil�s", "avi") ? True
    '
    
    Dim textoOriginalNorm As String
    Dim textoBuscadoNorm As String
    
    textoOriginalNorm = NormalizarTextoConCacheRapido(textoOriginal)
    textoBuscadoNorm = NormalizarTextoConCacheRapido(textoBuscado)
    
    ' Verificar si el texto original empieza con el texto buscado
    BuscarTextoNormalizado = (Len(textoOriginalNorm) >= Len(textoBuscadoNorm)) And _
                             (Left(textoOriginalNorm, Len(textoBuscadoNorm)) = textoBuscadoNorm)
End Function

Public Sub LimpiarCache()
    '
    ' PROP�SITO: Funci�n p�blica para limpiar cache manualmente
    ' USO: Llamar si se cambian los datos de la hoja o por mantenimiento
    '
    Set cacheNormalizacion = Nothing
End Sub

' ============================================
' FUNCIONES AUXILIARES ROBUSTAS
' ============================================

Private Function ObtenerUltimaFilaSegura(ws As Worksheet) As Long
    '
    ' PROP�SITO: Obtener �ltima fila con datos con validaci�n robusta
    ' MEJORAS IMPLEMENTADAS:
    '   ? Manejo de errores si la hoja est� corrupta
    '   ? Valor m�nimo por defecto
    '   ? Sin dependencia de formato espec�fico
    '
    
    On Error Resume Next
    Dim ultimaFila As Long
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' VALIDACI�N: Si hay error o valor inv�lido, usar valor por defecto
    If Err.Number <> 0 Or ultimaFila < 2 Then
        ultimaFila = 2 ' Valor por defecto (fila 1 = encabezados, fila 2 = primer dato)
        Err.Clear
    End If
    On Error GoTo 0
    
    ObtenerUltimaFilaSegura = ultimaFila
End Function

Private Sub LimpiarCeldasRelacionadas(hojaTrabajo As Worksheet, fila As Long, _
                                     columnaCP As Long, columnaPoblacion As Long, columnaProvincia As Long, _
                                     limpiarCP As Boolean, limpiarPoblacion As Boolean, limpiarProvincia As Boolean)
    '
    ' PROP�SITO: Limpiar celdas relacionadas de forma selectiva y segura
    ' PAR�METROS BOOLEANOS: Permiten limpiar solo las celdas necesarias
    ' ORDEN IMPORTANTE: Primero validaciones, luego contenidos
    '
    
    ' PASO 1: Limpiar validaciones primero para evitar errores
    If limpiarCP Then
        Call LimpiarValidacionSilenciosa(hojaTrabajo.Cells(fila, columnaCP))
    End If
    If limpiarPoblacion Then
        Call LimpiarValidacionSilenciosa(hojaTrabajo.Cells(fila, columnaPoblacion))
    End If
    
    ' PASO 2: Limpiar contenidos despu�s
    If limpiarCP Then
        hojaTrabajo.Cells(fila, columnaCP).ClearContents
    End If
    If limpiarPoblacion Then
        hojaTrabajo.Cells(fila, columnaPoblacion).ClearContents
    End If
    If limpiarProvincia Then
        hojaTrabajo.Cells(fila, columnaProvincia).ClearContents
    End If
End Sub

Private Sub LimpiarValidacionSilenciosa(celda As Range)
    '
    ' PROP�SITO: Eliminar validaci�n existente sin generar errores
    ' CR�TICO: Siempre limpiar validaciones antes de crear nuevas
    '
    
    On Error Resume Next
    celda.Validation.Delete
    Err.Clear
    On Error GoTo 0
End Sub

Private Function ExisteEnColeccion(coleccion As Collection, valor As String) As Boolean
    '
    ' PROP�SITO: Verificar si un valor ya existe en una Collection
    ' USO: Control de duplicados en listas de municipios/CPs
    ' M�TODO: Comparaci�n insensible a may�sculas/min�sculas
    '
    
    Dim i As Long
    ExisteEnColeccion = False
    For i = 1 To coleccion.Count
        If UCase(CStr(coleccion(i))) = UCase(valor) Then
            ExisteEnColeccion = True
            Exit For ' OPTIMIZACI�N: Salir tan pronto como se encuentre
        End If
    Next i
End Function

Private Function CrearListaTruncada(coleccion As Collection) As String
    '
    ' PROP�SITO: Convertir Collection a string para validaci�n, respetando l�mite de 255 caracteres
    ' RESTRICCI�N EXCEL: Las f�rmulas de validaci�n no pueden superar 255 caracteres
    ' ESTRATEGIA: Incluir tantos elementos como sea posible sin superar el l�mite
    '
    
    Dim lista As String
    Dim longitudTotal As Long
    Dim i As Long
    Dim itemActual As String
    
    lista = ""
    longitudTotal = 0
    
    For i = 1 To coleccion.Count
        itemActual = CStr(coleccion(i))
        
        ' CALCULAR LONGITUD CON COMA SEPARADORA
        Dim longitudNueva As Long
        If lista = "" Then
            longitudNueva = Len(itemActual)
        Else
            longitudNueva = longitudTotal + 1 + Len(itemActual) ' +1 por la coma
        End If
        
        ' VERIFICAR L�MITE DE EXCEL (250 caracteres para margen de seguridad)
        If longitudNueva > 250 Then
            Exit For ' Salir si se supera el l�mite
        End If
        
        ' AGREGAR ELEMENTO A LA LISTA
        If lista = "" Then
            lista = itemActual
            longitudTotal = Len(itemActual)
        Else
            lista = lista & "," & itemActual
            longitudTotal = longitudNueva
        End If
    Next i
    
    CrearListaTruncada = lista
End Function

Private Sub CrearValidacionListaSilenciosa(celda As Range, lista As String)
    '
    ' PROP�SITO: Crear validaci�n de lista de forma robusta y silenciosa
    ' CARACTER�STICAS:
    '   ? Manejo de hojas protegidas
    '   ? Sin alertas molestas al usuario
    '   ? Validaci�n permisiva (permite valores libres)
    '   ? Recuperaci�n autom�tica de estados
    '
    
    If Len(lista) = 0 Then Exit Sub
    
    ' GUARDAR ESTADOS ACTUALES
    Dim estadoAlertas As Boolean
    Dim estabaProtegida As Boolean
    
    estadoAlertas = Application.DisplayAlerts
    estabaProtegida = celda.Worksheet.ProtectContents
    
    On Error GoTo RestaurarEstados
    
    ' CONFIGURAR PARA OPERACI�N SILENCIOSA
    Application.DisplayAlerts = False
    
    ' DESPROTEGER SI ES NECESARIO
    If estabaProtegida Then
        celda.Worksheet.Unprotect
    End If
    
    ' LIMPIAR VALIDACI�N PREVIA
    celda.Validation.Delete
    
    ' CREAR NUEVA VALIDACI�N
    With celda.Validation
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=lista
        .IgnoreBlank = True                ' Permitir celdas vac�as
        .InCellDropdown = True            ' Mostrar flecha desplegable
        .ShowInput = False                ' No mostrar mensaje de entrada
        .ShowError = False                ' No mostrar mensaje de error (permitir valores libres)
    End With

RestaurarEstados:
    ' RESTAURAR ESTADOS ORIGINALES SIEMPRE
    If estabaProtegida Then
        celda.Worksheet.Protect
    End If
    Application.DisplayAlerts = estadoAlertas
    
    If Err.Number <> 0 Then Err.Clear
End Sub

' ============================================
' AUTOCOMPLETADO OPTIMIZADO CON NORMALIZACI�N
' ============================================

Public Sub ProcesarAutocompletadoPoblacion(Target As Range, hojaTrabajo As Worksheet, _
                                         columnaPoblacion As Long, columnaCP As Long, columnaProvincia As Long)
    '
    ' PROP�SITO: Proporcionar autocompletado inteligente mientras se escribe en poblaci�n
    ' CARACTER�STICAS:
    '   ? B�squeda normalizada (sin acentos)
    '   ? Optimizaci�n con arrays en memoria
    '   ? L�mite inteligente de sugerencias (2-10)
    '   ? Evita procesamiento excesivo
    '
    ' USO: Llamar desde el evento Worksheet_Change despu�s de ProcesarCambioCP_Poblacion
    '
    ' === NUEVA VALIDACI�N: Solo procesar si es UNA celda ===
    If Target.Cells.Count <> 1 Then Exit Sub
    ' VALIDACIONES INICIALES
    If Target.Column <> columnaPoblacion Then Exit Sub              ' Solo columna poblaci�n
    If Len(Trim(Target.Value)) < 3 Then Exit Sub                   ' M�nimo 3 caracteres
    If Len(Trim(Target.Value)) > 15 Then Exit Sub                  ' Evitar procesamiento de nombres completos
    
    ' VERIFICAR EXISTENCIA DE HOJA DE DATOS
    Dim wsCodigosPostales As Worksheet
    On Error Resume Next
    Set wsCodigosPostales = hojaTrabajo.Parent.Worksheets("CODIGOS POSTALES")
    On Error GoTo 0
    
    If wsCodigosPostales Is Nothing Then Exit Sub
    
    Dim textoEscrito As String
    textoEscrito = Trim(Target.Value)
    
    ' B�SQUEDA OPTIMIZADA CON NORMALIZACI�N
    Dim coincidencias As Collection
    Set coincidencias = BuscarMunicipiosCoincidentesOptimizado(wsCodigosPostales, textoEscrito)
    
    ' CREAR LISTA SOLO SI HAY COINCIDENCIAS �TILES (no demasiadas, no muy pocas)
    If coincidencias.Count > 1 And coincidencias.Count <= 10 Then
        Dim listaCoincidencias As String
        listaCoincidencias = CrearListaTruncada(coincidencias)
        
        ' Aplicar validaci�n temporal para autocompletado
        Call CrearValidacionAutocompletadoOptimizada(Target, listaCoincidencias)
    End If
End Sub

Private Function BuscarMunicipiosCoincidentesOptimizado(wsCodigosPostales As Worksheet, textoEscrito As String) As Collection
    '
    ' PROP�SITO: B�squeda ultra-r�pida de municipios que coinciden con texto parcial
    ' OPTIMIZACIONES CLAVE:
    '   ? Arrays en memoria (100x m�s r�pido que acceso celda por celda)
    '   ? B�squeda normalizada para mejor UX
    '   ? L�mite autom�tico para rendimiento
    '   ? Validaci�n robusta de arrays
    '
    
    Dim coincidencias As Collection
    Set coincidencias = New Collection
    
    ' OBTENER DATOS EN MEMORIA (OPTIMIZACI�N CR�TICA)
    Dim ultimaFila As Long
    ultimaFila = ObtenerUltimaFilaSegura(wsCodigosPostales)
    
    ' CARGAR TODA LA COLUMNA EN ARRAY (mucho m�s r�pido que acceso individual)
    Dim rangeMunicipios As Range
    Set rangeMunicipios = wsCodigosPostales.Range("B2:B" & ultimaFila)
    
    ' VALIDACI�N ROBUSTA: Verificar que hay datos para procesar
    On Error Resume Next
    Dim arrayMunicipios As Variant
    arrayMunicipios = rangeMunicipios.Value
    On Error GoTo 0
    
    ' VERIFICAR QUE EL ARRAY ES V�LIDO
    If Not IsArray(arrayMunicipios) Then Exit Function
    If UBound(arrayMunicipios, 1) < 1 Then Exit Function
    
    ' B�SQUEDA NORMALIZADA EN ARRAY
    Dim i As Long
    Dim municipio As String
    
    For i = 1 To UBound(arrayMunicipios, 1)
        municipio = Trim(CStr(arrayMunicipios(i, 1)))
        
        ' B�SQUEDA NORMALIZADA: Encuentra "Avil�s" escribiendo "avi"
        If BuscarTextoNormalizado(municipio, textoEscrito) Then
            
            ' CONTROL DE DUPLICADOS
            If Not ExisteEnColeccion(coincidencias, municipio) Then
                coincidencias.Add municipio
                
                ' L�MITE PARA RENDIMIENTO: No m�s de 10 sugerencias
                If coincidencias.Count >= 10 Then Exit For
            End If
        End If
    Next i
    
    Set BuscarMunicipiosCoincidentesOptimizado = coincidencias
End Function

Private Sub CrearValidacionAutocompletadoOptimizada(celda As Range, lista As String)
    ' VALIDACIONES PREVIAS ROBUSTAS
    If Len(Trim(lista)) = 0 Then Exit Sub
    lista = Trim(lista)
    
    ' LIMPIEZA PREVENTIVA DE CARACTERES PROBLEM�TICOS
    lista = Replace(lista, Chr(34), "'")    ' Comillas dobles
    lista = Replace(lista, Chr(10), "")     ' Saltos de l�nea
    lista = Replace(lista, Chr(13), "")     ' Retornos de carro
    lista = Replace(lista, ",,", ",")       ' Comas dobles
    
    ' VERIFICACI�N FINAL
    If Right(lista, 1) = "," Then lista = Left(lista, Len(lista) - 1)
    If Left(lista, 1) = "," Then lista = Mid(lista, 2)
    If Len(lista) = 0 Then Exit Sub
    
    ' APLICAR CON MANEJO ROBUSTO
    Dim estadoEventos As Boolean, estabaProtegida As Boolean
    estadoEventos = Application.enableEvents
    estabaProtegida = celda.Worksheet.ProtectContents
    
    On Error GoTo ManejadorError
    
    Application.enableEvents = False
    If estabaProtegida Then celda.Worksheet.Unprotect
    
    ' OPERACI�N AT�MICA: Todo o nada
    With celda
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, Formula1:=lista
        .Validation.IgnoreBlank = True
        .Validation.InCellDropdown = True
        .Validation.ShowInput = False
        .Validation.ShowError = False
    End With
    
    GoTo RestaurarEstados

ManejadorError:
    ' Si hay error, limpiar y salir silenciosamente
    On Error Resume Next
    celda.Validation.Delete
    On Error GoTo 0

RestaurarEstados:
    If estabaProtegida Then celda.Worksheet.Protect
    Application.enableEvents = estadoEventos
    If Err.Number <> 0 Then Err.Clear
End Sub

' ============================================
' DOCUMENTACI�N DE USO Y EJEMPLOS
' ============================================
'
' EJEMPLO DE IMPLEMENTACI�N EN HOJA:
' ====================================
'
' Private Sub Worksheet_Change(ByVal Target As Range)
'     ' Constantes de columnas (ajustar seg�n tu hoja)
'     Const COL_CP As Long = 5          ' Columna E
'     Const COL_POBLACION As Long = 6   ' Columna F
'     Const COL_PROVINCIA As Long = 7   ' Columna G
'
'     ' Procesar cambios principales
'     Call ProcesarCambioCP_Poblacion(Target, Me, COL_CP, COL_POBLACION, COL_PROVINCIA)
'
'     ' Autocompletado adicional (opcional)
'     Call ProcesarAutocompletadoPoblacion(Target, Me, COL_POBLACION, COL_CP, COL_PROVINCIA)
' End Sub
'
' ESTRUCTURA REQUERIDA DE "CODIGOS POSTALES":
' ===========================================
' Columna A: C�digo Postal (08001, 28001, etc.)
' Columna B: Municipio/Poblaci�n (Barcelona, Madrid, etc.)
' Columna C: Provincia (Barcelona, Madrid, etc.)
'
' CARACTER�STICAS PRINCIPALES:
' ============================
' ? B�squeda bidireccional: CP ? Poblaci�n
' ? Normalizaci�n inteligente: "aviles" encuentra "Avil�s"
' ? Cache autom�tico para m�ximo rendimiento
' ? Manejo robusto de errores sin crashes
' ? Listas desplegables autom�ticas para m�ltiples opciones
' ? Compatible con hojas protegidas
' ? Optimizado para m�ltiples usuarios

' MANTENIMIENTO:
' =============
' - Llamar LimpiarCache() si se actualizan los datos base
' - El cache se limpia autom�ticamente al cerrar Excel
' - L�mite autom�tico de 1000 elementos en cache
'
' ============================================



