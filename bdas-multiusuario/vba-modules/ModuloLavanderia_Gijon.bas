Attribute VB_Name = "ModuloLavanderia_Gijon"
Option Explicit

'==============================================================================
' M�DULO:  ModuloLavanderia_Gijon
'
' FUNCIONALIDADES PRINCIPALES:
' 1. Confirmar env�o a lavander�a Gij�n desde bot�n en hoja
' 2. Permite dos tipos de env�o:
'    - CON N� ORDEN: comportamiento normal del sistema
'    - SIN N� ORDEN: solo para estudios Est.1, Est.2, Est.3 (EXCEPCI�N ESPECIAL)
' 3. Registra UNA sola fila en hist�rico con totales consolidados
' 4. Excluye autom�ticamente duplicados ya enviados
' 5. Doble confirmaci�n antes de enviar (resumen + irreversibilidad)
'
' CARACTER�STICAS CLAVE:
' - Identificador alternativo para estudios: usa valor columna B (alojamiento)
' - No altera flujo existente para N� ORDEN normales
' - Mantiene consistencia en hist�rico
' - Compatible con todo el sistema Core
'==============================================================================

'---------------------------------------------------------------
' PROCEDIMIENTO PRINCIPAL: ConfirmarEnvioGijon
'
' OBJETIVO: Gestionar el proceso completo de env�o a lavander�a
' LLAMADA: Desde bot�n "CONFIRMAR ENV�O" en hoja "Lavander�a Gij�n"
'
' FLUJO GENERAL:
' 1. Preparaci�n y validaci�n inicial
' 2. Recopilar identificadores v�lidos (N� ORDEN o Estudios)
' 3. Verificar y excluir duplicados
' 4. Confirmaciones de usuario (2 pasos)
' 5. Registrar en hist�rico
' 6. Feedback al usuario
'---------------------------------------------------------------
Public Sub ConfirmarEnvioGijon()

    '===========================================================
    ' DECLARACI�N DE VARIABLES
    '===========================================================
    Dim hoja As Worksheet                    ' Hoja "Lavander�a Gij�n" de trabajo
    Dim ultimaFilaDatos As Long             ' �ltima fila con datos v�lidos (excluye TOTAL)
    Dim filaTotal As Long                   ' Fila donde est� la palabra "TOTAL"
    Dim fila As Long                        ' Contador para bucles
    Dim numOrden As Variant                 ' Valor de columna A (N� ORDEN)
    Dim alojamiento As String               ' Valor de columna B (ALOJAMIENTO)
    Dim identificador As String             ' Identificador a usar (N� ORDEN o Est.x)
    Dim listaOrdenesCompleta As String      ' Lista completa de identificadores (con duplicados)
    Dim listaOrdenesValidas As String       ' Lista depurada (sin duplicados)
    Dim ordenesDuplicadas As String         ' Identificadores ya enviados anteriormente
    Dim ordenesExcluidas As String          ' Identificadores excluidos del env�o actual
    Dim respuesta As VbMsgBoxResult         ' Respuesta del usuario en mensajes
    Dim cantidadOrdenesTotal As Long        ' Total de identificadores encontrados
    Dim cantidadOrdenesValidas As Long      ' Total de identificadores a enviar (sin duplicados)
    
    '===========================================================
    ' PASO 1: PREPARACI�N INICIAL
    '===========================================================
    
    ' Establecer referencia a la hoja de trabajo
    Set hoja = ThisWorkbook.Worksheets("Lavander�a Gij�n")
    
    ' Mantenimiento: Limpiar hist�rico antiguo (>3 meses) para optimizar rendimiento
    Call ModuloLavanderia_Core.LimpiarHistoricoAntiguo("HISTORICO_LAVANDERIA_GIJON")
    
    
    
    ' Determinar l�mites de datos:
    ' - �ltima fila con datos v�lidos (excluyendo fila TOTAL)
    ' - Fila donde est� "TOTAL" (para c�lculos posteriores)
    ultimaFilaDatos = ObtenerUltimaFilaDatos(hoja)
    filaTotal = BuscarFilaTotal(hoja)
    
    ' Validaci�n 1: Verificar que hay datos para procesar
    ' (Si no hay filas de datos o solo est� la cabecera)
    If ultimaFilaDatos < 2 Then
        MsgBox "No hay registros para enviar a lavander�a.", vbInformation
        Exit Sub
    End If
    
    ' Validaci�n 2: Verificar que existe fila TOTAL
    ' (Necesaria para c�lculos de sumatorias)
    If filaTotal = 0 Then
        MsgBox "No se encontr� la fila TOTAL.", vbExclamation
        Exit Sub
    End If
    
    '===========================================================
    ' PASO 2: RECOPILAR IDENTIFICADORES V�LIDOS
    '
    ' OBJETIVO: Construir lista de identificadores separados por comas
    '
    ' L�GICA DE SELECCI�N:
    ' 1. Si columna A tiene N� ORDEN ? usar N� ORDEN como identificador
    ' 2. Si columna A vac�a Y columna B es Est.1/Est.2/Est.3 ? usar Est.x como identificador
    ' 3. Si columna A vac�a Y columna B NO es estudio ? IGNORAR fila
    '
    ' IMPORTANTE: Esta es la MODIFICACI�N CLAVE que permite enviar estudios sin N� ORDEN
    '===========================================================
    listaOrdenesCompleta = ""
    cantidadOrdenesTotal = 0
    
    ' Recorrer todas las filas de datos (desde fila 2 hasta �ltima fila con datos)
    For fila = 2 To ultimaFilaDatos
    
        ' Leer valores de las columnas clave:
        numOrden = hoja.Cells(fila, "A").Value          ' N� ORDEN (columna A)
        alojamiento = Trim(CStr(hoja.Cells(fila, "B").Value))  ' ALOJAMIENTO (columna B, sin espacios)
        
        ' Excluir fila TOTAL (palabra "TOTAL" en cualquier may�scula/min�scula)
        If UCase(alojamiento) = "TOTAL" Then GoTo SiguienteFila
        
        ' *** L�GICA DE IDENTIFICACI�N (MODIFICADA) ***
        ' Caso A: Tiene N� ORDEN en columna A ? comportamiento normal
        If Trim(CStr(numOrden)) <> "" Then
            identificador = Trim(CStr(numOrden))
        
        ' Caso B: NO tiene N� ORDEN PERO es Estudio (Est.1, Est.2, Est.3)
        ' Esta es la EXCEPCI�N que permite enviar estudios sin N� ORDEN
        ElseIf alojamiento = "Est.1" Or alojamiento = "Est.2" Or alojamiento = "Est.3" Then
            identificador = alojamiento  ' Usar "Est.x" como identificador alternativo
        
        ' Caso C: No tiene N� ORDEN y NO es estudio ? EXCLUIR del env�o
        Else
            GoTo SiguienteFila  ' Saltar a siguiente fila
        End If
        
        ' A�adir identificador a la lista:
        ' - Si es el primer identificador: iniciar lista
        ' - Si ya hay identificadores: a�adir con separador ", "
        If listaOrdenesCompleta = "" Then
            listaOrdenesCompleta = identificador
        Else
            listaOrdenesCompleta = listaOrdenesCompleta & ", " & identificador
        End If
        
        ' Incrementar contador de identificadores v�lidos
        cantidadOrdenesTotal = cantidadOrdenesTotal + 1
        
SiguienteFila:
    Next fila
    
    ' Validaci�n 3: Verificar que se encontraron identificadores v�lidos
    If cantidadOrdenesTotal = 0 Then
        MsgBox "No hay identificadores v�lidos para enviar.", vbInformation
        Exit Sub
    End If
    
    '===========================================================
    ' PASO 3: VERIFICACI�N DE DUPLICADOS
    '
    ' OBJETIVO: Evitar reenv�os de identificadores ya registrados
    '
    ' FUNCIONALIDAD:
    ' 1. Consulta al m�dulo Core para verificar duplicados en hist�rico
    ' 2. Si hay duplicados ? mostrar advertencia y preguntar
    ' 3. Si usuario acepta ? excluir duplicados del env�o actual
    ' 4. Si usuario rechaza ? cancelar env�o
    '
    ' NOTA: Para estudios sin N� ORDEN, busca "Est.x" en hist�rico
    '===========================================================
    
    ' Consultar al m�dulo Core qu� identificadores ya existen en hist�rico
    ordenesDuplicadas = ModuloLavanderia_Core.ObtenerOrdenesDuplicadas( _
                          listaOrdenesCompleta, "HISTORICO_LAVANDERIA_GIJON")
    
    ' Si se encontraron duplicados...
    If ordenesDuplicadas <> "" Then
        ' Mostrar advertencia con lista de duplicados
        respuesta = MsgBox( _
            "Los siguientes identificadores ya fueron enviados:" & vbCrLf & vbCrLf & _
            ordenesDuplicadas & vbCrLf & vbCrLf & _
            "Ser�n excluidos del env�o." & vbCrLf & _
            "�Desea continuar?", _
            vbQuestion + vbYesNo)
        
        ' Si usuario rechaza ? cancelar todo el proceso
        If respuesta = vbNo Then Exit Sub
        
        ' Excluir duplicados de la lista (funci�n auxiliar local)
        listaOrdenesValidas = ExcluirOrdenesDuplicadas(listaOrdenesCompleta, ordenesDuplicadas)
        ordenesExcluidas = ordenesDuplicadas  ' Guardar para mostrar en resumen
    Else
        ' Si no hay duplicados ? usar lista completa
        listaOrdenesValidas = listaOrdenesCompleta
        ordenesExcluidas = ""
    End If
    
    ' Validaci�n 4: Verificar que quedan identificadores despu�s de excluir duplicados
    If Trim(listaOrdenesValidas) = "" Then
        MsgBox "No quedan registros nuevos para enviar.", vbInformation
        Exit Sub
    End If
    
    ' Contar cu�ntos identificadores se enviar�n realmente
    cantidadOrdenesValidas = ContarOrdenes(listaOrdenesValidas)
    
    '===========================================================
    ' PASO 4: CONFIRMACI�N DEL USUARIO
    '
    ' OBJETIVO: Validar que el usuario realmente quiere proceder
    '
    ' FLUJO DE CONFIRMACI�N:
    ' 1. Mostrar resumen del env�o (qu� se enviar�)
    ' 2. Usuario confirma o cancela
    '
    ' NOTA: Versi�n simplificada (original ten�a doble confirmaci�n)
    '===========================================================
    
    respuesta = MsgBox( _
        "Se enviar�n los siguientes identificadores:" & vbCrLf & _
        listaOrdenesValidas & vbCrLf & vbCrLf & _
        "Cantidad: " & cantidadOrdenesValidas & vbCrLf & _
        "�Desea continuar?", _
        vbQuestion + vbYesNo)
    
    ' Si usuario cancela ? terminar procedimiento
    If respuesta = vbNo Then Exit Sub
    
    '===========================================================
    ' PASO 5: REGISTRO EN HIST�RICO
    '
    ' OBJETIVO: Guardar datos consolidados en hoja hist�rico
    '
    ' CARACTER�STICAS:
    ' - Usa funci�n del m�dulo Core (mantiene consistencia)
    ' - Registra UNA sola fila con totales sumarizados
    ' - Maneja autom�ticamente: acumulaci�n si mismo d�a, creaci�n si nuevo d�a
    ' - Para estudios sin N� ORDEN: "Est.x" se almacena como N� ORDEN en hist�rico
    '===========================================================
    
'===========================================================
    ' PASO 5a: SOLICITAR FECHA DE ENV�O AL USUARIO
    '
    ' - Se muestra un InputBox con la fecha de hoy como valor por defecto
    ' - Se valida formato (DD/MM/AAAA) y que la fecha exista en el calendario
    ' - Si el usuario cancela ? se cancela el env�o
    ' - Si introduce una fecha incorrecta ? se avisa y se vuelve a preguntar
    ' - Usa funciones de validaci�n ya existentes en ModuloLavanderia_Core
    '===========================================================
    Dim fechaEnvio As Variant
    Dim textoFecha As String
    Dim fechaValida As Boolean
    
    fechaValida = False
    
    ' Bucle hasta obtener una fecha v�lida o que el usuario cancele
    Do While Not fechaValida
    
        ' Mostrar InputBox con fecha de hoy como valor por defecto
        textoFecha = InputBox( _
            "Introduzca la fecha de env�o a lavander�a:" & vbCrLf & vbCrLf & _
            "Formato: DD/MM/AAAA", _
            "Fecha de env�o", _
            Format(Date, "DD/MM/YYYY"))  ' Valor por defecto = hoy
        
        ' Si el usuario pulsa Cancelar ? textoFecha = "" ? cancelar env�o
        If textoFecha = "" Then
            MsgBox "Env�o cancelado.", vbInformation
            Exit Sub
        End If
        
        ' Validar formato DD/MM/AAAA o DD-MM-AAAA
        If Not ModuloLavanderia_Core.EsFormatoFechaValido(textoFecha) Then
            MsgBox "Formato de fecha incorrecto." & vbCrLf & vbCrLf & _
                   "Introduzca la fecha con el formato DD/MM/AAAA.", _
                   vbExclamation, "Fecha incorrecta"
            GoTo SiguienteIntento  ' Volver a pedir fecha
        End If
        
        ' Validar que la fecha exista en el calendario (ej: no 31/02/2025)
        If Not ModuloLavanderia_Core.EsFechaValida(textoFecha) Then
            MsgBox "La fecha introducida no existe en el calendario." & vbCrLf & vbCrLf & _
                   "Compruebe el d�a y el mes.", _
                   vbExclamation, "Fecha incorrecta"
            GoTo SiguienteIntento  ' Volver a pedir fecha
        End If
        
        ' Fecha v�lida: convertir texto a fecha y salir del bucle
        fechaEnvio = ModuloLavanderia_Core.ConvertirTextoAFecha(textoFecha)
        fechaValida = True
        
SiguienteIntento:
    Loop
    
    ' LLAMADA PRINCIPAL al sistema Core para registrar env�o
    ' Par�metros:
    ' 1. hoja: Hoja origen con los datos
    ' 2. filaTotal: Fila donde est�n los totales a sumarizar
    ' 3. listaOrdenesValidas: Identificadores a enviar (N� ORDEN o Est.x)
    ' 4. "HISTORICO_LAVANDERIA_GIJON": Nombre hoja hist�rico espec�fica
    ' 5. fechaEnvio: Fecha a registrar
    Call ModuloLavanderia_Core.RegistrarEnvioEnHistorico( _
            hoja, filaTotal, listaOrdenesValidas, _
            "HISTORICO_LAVANDERIA_GIJON", fechaEnvio)
    
'===========================================================
' PASO 5.5: ELIMINAR FILAS ENVIADAS CORRECTAMENTE
'
' Regla:
' - Si tiene N� ORDEN ? se elimina si est� en listaOrdenesValidas
' - Si NO tiene N� ORDEN y es Est.x ? se elimina si Est.x est� en la lista
'
' Se recorre DE ABAJO A ARRIBA para evitar saltos de filas
'===========================================================

Dim arrIds() As String
Dim idFila As String

arrIds = Split(listaOrdenesValidas, ",")

For fila = ultimaFilaDatos To 2 Step -1

    alojamiento = Trim(CStr(hoja.Cells(fila, "B").Value))
    numOrden = Trim(CStr(hoja.Cells(fila, "A").Value))
    
    ' Ignorar fila TOTAL por seguridad
    If UCase(alojamiento) = "TOTAL" Then GoTo SiguienteBorrado
    
    ' Determinar identificador real de la fila
    If numOrden <> "" Then
        idFila = numOrden
    ElseIf alojamiento = "Est.1" Or alojamiento = "Est.2" Or alojamiento = "Est.3" Then
        idFila = alojamiento
    Else
        GoTo SiguienteBorrado
    End If
    
    ' Si el identificador est� en la lista enviada ? borrar fila
    If IdentificadorEnLista(idFila, arrIds) Then
        hoja.Rows(fila).Delete
    End If

SiguienteBorrado:
Next fila

    
    '===========================================================
    ' PASO 6: FEEDBACK FINAL
    '===========================================================
    MsgBox "Env�o registrado correctamente.", vbInformation

End Sub

'===============================================================
' FUNCIONES AUXILIARES LOCALES
' (Solo usadas dentro de este m�dulo)
'===============================================================

'---------------------------------------------------------------
' FUNCI�N: ObtenerUltimaFilaDatos
'
' OBJETIVO: Determinar la �ltima fila con datos v�lidos en la hoja
'
' L�GICA MODIFICADA:
' - Considera v�lida una fila si:
'   1. Tiene N� ORDEN en columna A (comportamiento normal) O
'   2. Es Estudio (Est.1, Est.2, Est.3) aunque columna A vac�a
' - Excluye: filas TOTAL y filas completamente vac�as
'
' PAR�METROS:
'   hoja: Worksheet - Hoja a analizar
'
' RETORNO: Long - N�mero de �ltima fila con datos (1 si no hay)
'---------------------------------------------------------------
Private Function ObtenerUltimaFilaDatos(ByVal hoja As Worksheet) As Long
    Dim fila As Long
    Dim ultimaFila As Long
    Dim alojamiento As String
    
    ' Determinar l�mite de b�squeda: m�ximo entre �ltima fila con datos en A o B
    ultimaFila = Application.WorksheetFunction.Max( _
        hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row, _
        hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row)
    
    ' Recorrer de abajo hacia arriba (para encontrar r�pido la �ltima v�lida)
    For fila = ultimaFila To 2 Step -1
        
        ' Leer alojamiento de columna B
        alojamiento = Trim(CStr(hoja.Cells(fila, "B").Value))
        
        ' Excluir fila TOTAL
        If UCase(alojamiento) = "TOTAL" Then GoTo SiguienteFila
        
        ' Excluir filas completamente vac�as (A a U)
        If Application.WorksheetFunction.CountA(hoja.Range("A" & fila & ":U" & fila)) = 0 Then
            GoTo SiguienteFila
        End If
        
        ' *** L�GICA MODIFICADA: criterio de validez extendido ***
        ' Una fila es v�lida si:
        ' A) Tiene N� ORDEN en columna A (comportamiento normal) O
        ' B) Es Estudio (Est.1, Est.2, Est.3) aunque columna A vac�a
        If Trim(CStr(hoja.Cells(fila, "A").Value)) <> "" _
           Or alojamiento = "Est.1" _
           Or alojamiento = "Est.2" _
           Or alojamiento = "Est.3" Then
            
            ' Retornar esta fila como �ltima v�lida
            ObtenerUltimaFilaDatos = fila
            Exit Function
        End If
        
SiguienteFila:
    Next fila
    
    ' Si no encontr� filas v�lidas ? retornar 1 (solo cabecera)
    ObtenerUltimaFilaDatos = 1
End Function

'---------------------------------------------------------------
' FUNCI�N: BuscarFilaTotal
'
' OBJETIVO: Encontrar fila donde est� la palabra "TOTAL"
'
' BUSQUEDA: Desde fila 2 hasta 10 filas despu�s de �ltima fila con datos en B
'
' PAR�METROS:
'   hoja: Worksheet - Hoja a buscar
'
' RETORNO: Long - N�mero de fila donde est� TOTAL, 0 si no se encuentra
'---------------------------------------------------------------
Private Function BuscarFilaTotal(ByVal hoja As Worksheet) As Long
    Dim fila As Long
    
    ' Buscar en rango ampliado (hasta 10 filas despu�s de �ltima con datos)
    For fila = 2 To hoja.Cells(hoja.Rows.Count, "B").End(xlUp).Row + 10
        ' Buscar "TOTAL" en columna B (case insensitive)
        If UCase(Trim(hoja.Cells(fila, "B").Value)) = "TOTAL" Then
            BuscarFilaTotal = fila
            Exit Function
        End If
    Next fila
    
    ' Si no se encontr� ? retornar 0
    BuscarFilaTotal = 0
End Function

'---------------------------------------------------------------
' FUNCI�N: ContarOrdenes
'
' OBJETIVO: Contar cantidad de identificadores en lista separada por comas
'
' EJEMPLO: "1234, Est.1, 5678" ? retorna 3
'
' PAR�METROS:
'   listaOrdenes: String - Lista separada por comas
'
' RETORNO: Long - Cantidad de elementos
'---------------------------------------------------------------
Private Function ContarOrdenes(ByVal listaOrdenes As String) As Long
    ' Si lista vac�a ? retornar 0
    If Trim(listaOrdenes) = "" Then
        ContarOrdenes = 0
    Else
        ' Usar Split para dividir por comas y contar elementos
        ' UBound devuelve �ndice �ltimo elemento, +1 para cantidad
        ContarOrdenes = UBound(Split(listaOrdenes, ",")) + 1
    End If
End Function

'---------------------------------------------------------------
' FUNCI�N: ExcluirOrdenesDuplicadas
'
' OBJETIVO: Filtrar lista completa excluyendo elementos duplicados
'
' L�GICA: Comparaci�n exacta (case-sensitive) entre elementos
'
' PAR�METROS:
'   listaCompleta: String - Lista original con todos los elementos
'   listaDuplicadas: String - Lista de elementos a excluir
'
' RETORNO: String - Nueva lista sin los elementos duplicados
'---------------------------------------------------------------
Private Function ExcluirOrdenesDuplicadas(ByVal listaCompleta As String, _
                                         ByVal listaDuplicadas As String) As String
    Dim arrAll, arrDup, i As Long, j As Long
    Dim esDuplicada As Boolean
    Dim resultado As String
    
    ' Convertir listas de texto a arrays para procesamiento
    arrAll = Split(listaCompleta, ",")   ' Array con todos los elementos
    arrDup = Split(listaDuplicadas, ",") ' Array con duplicados a excluir
    
    ' Recorrer todos los elementos de la lista completa
    For i = LBound(arrAll) To UBound(arrAll)
        esDuplicada = False
        
        ' Comparar con cada elemento de la lista de duplicados
        For j = LBound(arrDup) To UBound(arrDup)
            ' Comparaci�n exacta (con trim para eliminar espacios)
            If Trim(arrAll(i)) = Trim(arrDup(j)) Then
                esDuplicada = True  ' Marcar como duplicado
                Exit For            ' Salir del bucle interno
            End If
        Next j
        
        ' Si NO es duplicado ? a�adir a resultado
        If Not esDuplicada Then
            If resultado = "" Then
                resultado = Trim(arrAll(i))
            Else
                resultado = resultado & ", " & Trim(arrAll(i))
            End If
        End If
    Next i
    
    ' Retornar lista filtrada
    ExcluirOrdenesDuplicadas = resultado
End Function

'---------------------------------------------------------------
' FUNCI�N: IdentificadorEnLista
'
' OBJETIVO:
' Verificar si un identificador concreto existe dentro de una
' lista de identificadores previamente enviados.
'
' USO PRINCIPAL:
' - Se utiliza en el proceso de eliminaci�n de filas tras env�o
' - Permite comprobar si una fila debe borrarse porque:
'   � Su N� ORDEN fue enviado, o
'   � Es un Estudio (Est.1 / Est.2 / Est.3) enviado sin N� ORDEN
'
' L�GICA:
' - Recorre un array de identificadores (Split de listaOrdenesValidas)
' - Compara cada elemento con el identificador de la fila
' - Comparaci�n exacta tras aplicar Trim (robusta ante espacios)
'
' PAR�METROS:
'   id  : String   - Identificador de la fila (N� ORDEN o Est.x)
'   arr : Variant  - Array de identificadores v�lidos enviados
'
' RETORNO:
'   Boolean - True si el identificador existe en la lista
'             False en caso contrario
'---------------------------------------------------------------
Private Function IdentificadorEnLista(ByVal id As String, _
                                      ByVal arr As Variant) As Boolean
    Dim i As Long
    
    ' Recorrer todos los identificadores enviados
    For i = LBound(arr) To UBound(arr)
        
        ' Comparaci�n exacta (Trim para evitar errores por espacios)
        If Trim(arr(i)) = Trim(id) Then
            IdentificadorEnLista = True
            Exit Function
        End If
    Next i
    
    ' Si no se encontr� coincidencia
    IdentificadorEnLista = False
End Function


'==============================================================================
' NOTAS IMPORTANTES DE IMPLEMENTACI�N:
'
' 1. COMPATIBILIDAD CON SISTEMA CORE:
'    - Los estudios sin N� ORDEN se manejan como "identificadores alternativos"
'    - El hist�rico almacenar� "Est.x" en columna de N� ORDEN
'    - Todas las funciones Core funcionan igual (buscando en columna A)
'
' 2. SEGURIDAD Y VALIDACI�N:
'    - Solo estudios espec�ficos (Est.1, Est.2, Est.3) pueden enviarse sin N� ORDEN
'    - Habitaciones normales (1, 2, 3, etc.) REQUIEREN N� ORDEN obligatorio
'    - Verificaci�n de duplicados aplica igual para ambos tipos
'
' 3. MANTENIMIENTO:
'    - Limpieza autom�tica de hist�rico (>3 meses) al inicio
'    - C�digo modular con funciones espec�ficas
'    - Comentarios detallados para futuras modificaciones
'
' 4. ESCALABILIDAD:
'    - F�cil agregar m�s estudios sin N� ORDEN modificando condiciones
'    - Compatible con otras lavander�as (cada una tiene su hist�rico)
'==============================================================================

