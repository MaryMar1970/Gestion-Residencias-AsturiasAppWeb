Attribute VB_Name = "CopiarDatosHojasGijonSotoOviedo"
'-------------------------------------------------------------------------------------------
' Macro: CopiaDatosHojasGijonSotoOviedo
'-------------------------------------------------------------------------------------------
' Descripci�n general:
'   - Copia datos desde un libro origen hacia el libro actual (destino).
'   - Permite seleccionar hojas: RESIDENCIA GIJ�N, RESIDENCIA SOTO, RESIDENCIA OVIEDO.
'   - Copia columnas coincidentes por nombre, siempre que no tengan f�rmulas en la fila 2
'     (tanto en origen como en destino), para evitar romper c�lculos.
'   - Optimiza rendimiento desactivando eventos/pantalla/c�lculo y copiando por bloques.
'   - Flujo especial para RESIDENCIA SOTO:
'       * Toma la columna "COMISI�N" del origen y la vuelca en "TURNO" del destino.
'       * Sustituye "NO" por "SIN TURNO" antes de escribir.
'       * Este paso se ejecuta primero para evitar sobrescrituras posteriores.
'   - NUEVO: Si la hoja destino es "RESIDENCIA GIJ�N", en el libro origen se acepta que
'     se llame "RESIDENCIA GIJ�N" o "RESIDENCIA UNIDAD".
'
' Uso:
'   1) Ejecuta la macro.
'   2) Selecciona el archivo origen.
'   3) Indica (por n�mero) qu� hojas copiar.
'   4) Se mostrar� un resumen con el resultado por hoja.
'-------------------------------------------------------------------------------------------
Sub CopiaDatosHojasGijonSotoOviedo()

    '------------------------------- Declaraciones ------------------------------------------
    Dim wbOrigen As Workbook, wbDestino As Workbook
    Dim wsOrigen As Worksheet, wsDestino As Worksheet
    Dim Ruta As Variant
    Dim candidatas As Collection, hojasSeleccionadas As Collection
    Dim listaHojas As String, respuesta As String
    Dim i As Integer, idx As Integer
    Dim encabezadosOrigen As Object, encabezadosDestino As Object
    Dim iColO As Long, iColD As Long
    Dim lastRowOrigen As Long, lastColOrigen As Long
    Dim tieneFormulaOrigen As Boolean, tieneFormulaDestino As Boolean
    Dim resultadoHojas As Object, msgResumen As String
    Dim arrNros() As String, nro As String
    Dim columnasCopiadas As Long
    Dim hojaNombre As Variant
    Dim colsToCopy As Collection, enc As Variant
    Dim colPair As Variant
    Dim colH_Origen As Long, colH_Destino As Long
    Dim esSoto As Boolean
    ' === NUEVAS VARIABLES PARA FILTRADO POR RANGO ===
    Dim criterioRespuesta As String
    Dim criterioCol As String
    Dim valorDesde As Variant, valorHasta As Variant
    Dim usarFiltro As Boolean
    Dim esFechaFiltro As Boolean
    Dim filaOrigen As Long
    Dim valorCelda As Variant
    Dim cumpleFiltro As Boolean
    Dim filasCopiadas As Long
    Dim primeraFilaRango As Long, ultimaFilaRango As Long
    ' Variables adicionales para filtro m�ltiple
    Dim usarFiltroMultiple As Boolean
    Dim criterioCol2 As String
    Dim valorDesde2 As Variant, valorHasta2 As Variant
    Dim valorCelda2 As Variant
    Dim cumpleFiltro2 As Boolean
    Dim usarFiltroResolucion As Boolean

    '--------------------------- Optimizaci�n de Excel --------------------------------------
    ' Desactivar eventos, actualizaci�n de pantalla y c�lculo autom�tico acelera la macro
    Application.enableEvents = False
    Application.screenUpdating = False
    Application.calculation = xlCalculationManual

    '--------------------------- Selecci�n de archivo origen --------------------------------
    ' El usuario elige el libro desde el que se copiar�n los datos
    Ruta = Application.GetOpenFilename( _
        "Archivos de Excel (*.xls; *.xlsx; *.xlsm), *.xls; *.xlsx; *.xlsm", , _
        "Selecciona el archivo origen" _
    )
    If Ruta = False Then GoTo Finalizar ' Cancelado por el usuario

    ' Libro destino = el actual; libro origen = el seleccionado (s�lo lectura)
    Set wbDestino = ThisWorkbook
    Set wbOrigen = Workbooks.Open(Ruta, ReadOnly:=True)

    '--------------------------- Definici�n de hojas disponibles -----------------------------
    Set candidatas = New Collection
    candidatas.Add "RESIDENCIA GIJ�N"
    candidatas.Add "RESIDENCIA SOTO"
    candidatas.Add "RESIDENCIA OVIEDO"

    ' Mostrar men� numerado para el InputBox
    listaHojas = ""
    For i = 1 To candidatas.Count
        listaHojas = listaHojas & i & ". " & candidatas(i) & vbCrLf
    Next i

    '--------------------------- Selecci�n de hojas por n�mero -------------------------------
    respuesta = InputBox( _
        "Selecciona los n�meros de las hojas a copiar, separados por coma:" & vbCrLf & vbCrLf & listaHojas, _
        "Seleccionar hojas a copiar" _
    )
    If Trim(respuesta) = "" Then
        MsgBox "No se ha seleccionado ninguna hoja. El proceso se cancela.", vbExclamation
        wbOrigen.Close SaveChanges:=False
        GoTo Finalizar
    End If

    ' Parseo de la respuesta y validaci�n de �ndices
    arrNros = Split(respuesta, ",")
    Set hojasSeleccionadas = New Collection
    For idx = LBound(arrNros) To UBound(arrNros)
        nro = Trim(arrNros(idx))
        If IsNumeric(nro) And val(nro) >= 1 And val(nro) <= candidatas.Count Then
            hojasSeleccionadas.Add candidatas(val(nro))
        End If
    Next idx

    If hojasSeleccionadas.Count = 0 Then
        MsgBox "No se ha seleccionado ning�n n�mero v�lido de hoja. El proceso se cancela.", vbExclamation
        wbOrigen.Close SaveChanges:=False
        GoTo Finalizar
    End If

    ' Diccionario para ir registrando el resultado de cada hoja
    Set resultadoHojas = CreateObject("Scripting.Dictionary")
    
    
            ' ========================== SELECCI�N DE CRITERIO DE FILTRADO ===========================
    criterioRespuesta = InputBox( _
        "Seleccione el criterio para filtrar las filas a copiar:" & vbCrLf & vbCrLf & _
        "1. N� ORDEN (columna A)" & vbCrLf & _
        "2. FECHA PETICI�N (columna B)" & vbCrLf & _
        "3. ENTRADA (columna L)" & vbCrLf & _
        "4. SALIDA (columna M)" & vbCrLf & _
        "5. ENTRADA + SALIDA (columnas L y M)" & vbCrLf & _
        "6. ENTRADA + RESOLUCI�N (columnas L y P)" & vbCrLf & _
        "7. SALIDA + RESOLUCI�N (columnas M y P)" & vbCrLf & _
        "8. Copiar TODAS las filas (sin filtro)" & vbCrLf & vbCrLf & _
        "Introduzca el n�mero (1-8):", _
        "Seleccionar criterio de filtrado", "8")
    
    ' Inicializar variables
    usarFiltro = False
    usarFiltroMultiple = False
    esFechaFiltro = False
    criterioCol = ""
    criterioCol2 = ""
    
    ' Variable para indicar si el segundo filtro es RESOLUCI�N (valores fijos)
    usarFiltroResolucion = False
    
    Select Case val(Trim(criterioRespuesta))
        Case 1
            criterioCol = "A"
            usarFiltro = True
            esFechaFiltro = False
        Case 2
            criterioCol = "B"
            usarFiltro = True
            esFechaFiltro = True
        Case 3
            criterioCol = "L"
            usarFiltro = True
            esFechaFiltro = True
        Case 4
            criterioCol = "M"
            usarFiltro = True
            esFechaFiltro = True
        Case 5
            ' Filtro m�ltiple:  ENTRADA + SALIDA
            criterioCol = "L"
            criterioCol2 = "M"
            usarFiltro = True
            usarFiltroMultiple = True
            esFechaFiltro = True
        Case 6
            ' Filtro m�ltiple:  ENTRADA + RESOLUCI�N
            criterioCol = "L"
            criterioCol2 = "P"
            usarFiltro = True
            usarFiltroMultiple = True
            usarFiltroResolucion = True
            esFechaFiltro = True
        Case 7
            ' Filtro m�ltiple:  SALIDA + RESOLUCI�N
            criterioCol = "M"
            criterioCol2 = "P"
            usarFiltro = True
            usarFiltroMultiple = True
            usarFiltroResolucion = True
            esFechaFiltro = True
        Case 8
            usarFiltro = False
        Case Else
            usarFiltro = False
    End Select
    
    ' Si hay filtro simple (no m�ltiple), pedir valores DESDE y HASTA
    If usarFiltro And Not usarFiltroMultiple Then
        If esFechaFiltro Then
            valorDesde = InputBox( _
                "Introduzca la FECHA DESDE (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                "Ejemplo: 01/01/2025", _
                "Fecha inicial del rango")
            If Trim(CStr(valorDesde)) = "" Then
                MsgBox "No se ha indicado fecha inicial.  Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
            ElseIf Not IsDate(valorDesde) Then
                MsgBox "El valor '" & valorDesde & "' no es una fecha v�lida.  Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
            Else
                valorDesde = CDate(valorDesde)
            End If
            
            If usarFiltro Then
                valorHasta = InputBox( _
                    "Introduzca la FECHA HASTA (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                    "Ejemplo: 31/12/2025", _
                    "Fecha final del rango")
                If Trim(CStr(valorHasta)) = "" Then
                    MsgBox "No se ha indicado fecha final. Se cancelar� el filtrado.", vbExclamation
                    usarFiltro = False
                ElseIf Not IsDate(valorHasta) Then
                    MsgBox "El valor '" & valorHasta & "' no es una fecha v�lida. Se cancelar� el filtrado.", vbExclamation
                    usarFiltro = False
                Else
                    valorHasta = CDate(valorHasta)
                End If
            End If
        Else
            ' Filtro num�rico (N� ORDEN)
            valorDesde = InputBox( _
                "Introduzca el N� ORDEN DESDE:" & vbCrLf & vbCrLf & _
                "Ejemplo: 1000", _
                "N� Orden inicial del rango")
            If Trim(CStr(valorDesde)) = "" Then
                MsgBox "No se ha indicado valor inicial. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
            ElseIf Not IsNumeric(valorDesde) Then
                MsgBox "El valor '" & valorDesde & "' no es num�rico. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
            Else
                valorDesde = CDbl(valorDesde)
            End If
            
            If usarFiltro Then
                valorHasta = InputBox( _
                    "Introduzca el N� ORDEN HASTA:" & vbCrLf & vbCrLf & _
                    "Ejemplo: 2000", _
                    "N� Orden final del rango")
                If Trim(CStr(valorHasta)) = "" Then
                    MsgBox "No se ha indicado valor final. Se cancelar� el filtrado.", vbExclamation
                    usarFiltro = False
                ElseIf Not IsNumeric(valorHasta) Then
                    MsgBox "El valor '" & valorHasta & "' no es num�rico. Se cancelar� el filtrado.", vbExclamation
                    usarFiltro = False
                Else
                    valorHasta = CDbl(valorHasta)
                End If
            End If
        End If
        
        ' Confirmar el rango seleccionado (filtro simple)
        If usarFiltro Then
            If esFechaFiltro Then
                MsgBox "Se copiar�n las filas donde " & vbCrLf & _
                       "la columna " & criterioCol & " est� entre:" & vbCrLf & vbCrLf & _
                       Format(valorDesde, "dd/mm/yyyy") & " y " & Format(valorHasta, "dd/mm/yyyy"), _
                       vbInformation, "Filtro configurado"
            Else
                MsgBox "Se copiar�n las filas donde " & vbCrLf & _
                       "la columna " & criterioCol & " (N� ORDEN) est� entre:" & vbCrLf & vbCrLf & _
                       valorDesde & " y " & valorHasta, _
                       vbInformation, "Filtro configurado"
            End If
        End If
    End If
    
    ' Si hay filtro M�LTIPLE con RESOLUCI�N (ENTRADA+RESOLUCI�N o SALIDA+RESOLUCI�N)
    If usarFiltro And usarFiltroMultiple And usarFiltroResolucion Then
        Dim nombreColumnaFecha As String
        If criterioCol = "L" Then
            nombreColumnaFecha = "ENTRADA"
        Else
            nombreColumnaFecha = "SALIDA"
        End If
        
        ' Pedir rango para la columna de fecha (L o M)
        valorDesde = InputBox( _
            "FILTRO " & nombreColumnaFecha & " (columna " & criterioCol & ")" & vbCrLf & vbCrLf & _
            "Introduzca la FECHA " & nombreColumnaFecha & " DESDE (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
            "Ejemplo: 01/01/2025", _
            "Fecha " & nombreColumnaFecha & " inicial")
        If Trim(CStr(valorDesde)) = "" Then
            MsgBox "No se ha indicado fecha inicial. Se cancelar� el filtrado.", vbExclamation
            usarFiltro = False
            usarFiltroMultiple = False
        ElseIf Not IsDate(valorDesde) Then
            MsgBox "El valor '" & valorDesde & "' no es una fecha v�lida. Se cancelar� el filtrado.", vbExclamation
            usarFiltro = False
            usarFiltroMultiple = False
        Else
            valorDesde = CDate(valorDesde)
        End If
        
        If usarFiltro Then
            valorHasta = InputBox( _
                "FILTRO " & nombreColumnaFecha & " (columna " & criterioCol & ")" & vbCrLf & vbCrLf & _
                "Introduzca la FECHA " & nombreColumnaFecha & " HASTA (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                "Ejemplo: 31/12/2025", _
                "Fecha " & nombreColumnaFecha & " final")
            If Trim(CStr(valorHasta)) = "" Then
                MsgBox "No se ha indicado fecha final. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            ElseIf Not IsDate(valorHasta) Then
                MsgBox "El valor '" & valorHasta & "' no es una fecha v�lida. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            Else
                valorHasta = CDate(valorHasta)
            End If
        End If
        
        ' Confirmar el filtro m�ltiple con RESOLUCI�N
        If usarFiltro And usarFiltroMultiple Then
            MsgBox "Se copiar�n las filas donde:" & vbCrLf & vbCrLf & _
                   nombreColumnaFecha & " (col " & criterioCol & ") entre:  " & Format(valorDesde, "dd/mm/yyyy") & " y " & Format(valorHasta, "dd/mm/yyyy") & vbCrLf & _
                   "RESOLUCI�N (col P) sea:  SI, CONCEDIDA o REEVALUADA" & vbCrLf & vbCrLf & _
                   "(Ambas condiciones deben cumplirse)", _
                   vbInformation, "Filtro " & nombreColumnaFecha & " + RESOLUCI�N configurado"
        End If
    End If
    
    ' Si hay filtro M�LTIPLE ENTRADA + SALIDA (sin RESOLUCI�N)
    If usarFiltro And usarFiltroMultiple And Not usarFiltroResolucion Then
        ' Pedir rango para ENTRADA (columna L)
        valorDesde = InputBox( _
            "FILTRO ENTRADA (columna L)" & vbCrLf & vbCrLf & _
            "Introduzca la FECHA ENTRADA DESDE (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
            "Ejemplo:  01/01/2025", _
            "Fecha ENTRADA inicial")
        If Trim(CStr(valorDesde)) = "" Then
            MsgBox "No se ha indicado fecha inicial de ENTRADA. Se cancelar� el filtrado.", vbExclamation
            usarFiltro = False
            usarFiltroMultiple = False
        ElseIf Not IsDate(valorDesde) Then
            MsgBox "El valor '" & valorDesde & "' no es una fecha v�lida.  Se cancelar� el filtrado.", vbExclamation
            usarFiltro = False
            usarFiltroMultiple = False
        Else
            valorDesde = CDate(valorDesde)
        End If
        
        If usarFiltro Then
            valorHasta = InputBox( _
                "FILTRO ENTRADA (columna L)" & vbCrLf & vbCrLf & _
                "Introduzca la FECHA ENTRADA HASTA (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                "Ejemplo: 31/12/2025", _
                "Fecha ENTRADA final")
            If Trim(CStr(valorHasta)) = "" Then
                MsgBox "No se ha indicado fecha final de ENTRADA. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            ElseIf Not IsDate(valorHasta) Then
                MsgBox "El valor '" & valorHasta & "' no es una fecha v�lida. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            Else
                valorHasta = CDate(valorHasta)
            End If
        End If
        
        ' Pedir rango para SALIDA (columna M)
        If usarFiltro Then
            valorDesde2 = InputBox( _
                "FILTRO SALIDA (columna M)" & vbCrLf & vbCrLf & _
                "Introduzca la FECHA SALIDA DESDE (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                "Ejemplo: 01/01/2025", _
                "Fecha SALIDA inicial")
            If Trim(CStr(valorDesde2)) = "" Then
                MsgBox "No se ha indicado fecha inicial de SALIDA. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            ElseIf Not IsDate(valorDesde2) Then
                MsgBox "El valor '" & valorDesde2 & "' no es una fecha v�lida. Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            Else
                valorDesde2 = CDate(valorDesde2)
            End If
        End If
        
        If usarFiltro Then
            valorHasta2 = InputBox( _
                "FILTRO SALIDA (columna M)" & vbCrLf & vbCrLf & _
                "Introduzca la FECHA SALIDA HASTA (formato dd/mm/aaaa):" & vbCrLf & vbCrLf & _
                "Ejemplo: 31/12/2025", _
                "Fecha SALIDA final")
            If Trim(CStr(valorHasta2)) = "" Then
                MsgBox "No se ha indicado fecha final de SALIDA.  Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            ElseIf Not IsDate(valorHasta2) Then
                MsgBox "El valor '" & valorHasta2 & "' no es una fecha v�lida.  Se cancelar� el filtrado.", vbExclamation
                usarFiltro = False
                usarFiltroMultiple = False
            Else
                valorHasta2 = CDate(valorHasta2)
            End If
        End If
        
        ' Confirmar el rango seleccionado (filtro m�ltiple ENTRADA+SALIDA)
        If usarFiltro And usarFiltroMultiple Then
            MsgBox "Se copiar�n las filas donde:" & vbCrLf & vbCrLf & _
                   "ENTRADA (col L) entre: " & Format(valorDesde, "dd/mm/yyyy") & " y " & Format(valorHasta, "dd/mm/yyyy") & vbCrLf & _
                   "SALIDA (col M) entre: " & Format(valorDesde2, "dd/mm/yyyy") & " y " & Format(valorHasta2, "dd/mm/yyyy") & vbCrLf & vbCrLf & _
                   "(Ambas condiciones deben cumplirse)", _
                   vbInformation, "Filtro m�ltiple configurado"
        End If
    End If



    '=========================== Procesamiento por hoja seleccionada =========================
    For Each hojaNombre In hojasSeleccionadas

        Set wsOrigen = Nothing
        Set wsDestino = Nothing

        '----------------------- Selecci�n robusta de hoja origen/destino --------------------
        ' Destino: siempre el nombre elegido
        On Error Resume Next
        Set wsDestino = wbDestino.Worksheets(hojaNombre)
        On Error GoTo 0

        ' Origen:
        ' - Si destino es "RESIDENCIA GIJ�N": admitir que el origen se llame "RESIDENCIA GIJ�N" o "RESIDENCIA UNIDAD"
        ' - Si no: usar el mismo nombre
        If hojaNombre = "RESIDENCIA GIJ�N" Then
            On Error Resume Next
            Set wsOrigen = wbOrigen.Worksheets("RESIDENCIA GIJ�N")
            If wsOrigen Is Nothing Then
                Set wsOrigen = wbOrigen.Worksheets("RESIDENCIA UNIDAD")
            End If
            On Error GoTo 0
        Else
            On Error Resume Next
            Set wsOrigen = wbOrigen.Worksheets(hojaNombre)
            On Error GoTo 0
        End If

        ' Validaciones de existencia
        If wsOrigen Is Nothing Then
            resultadoHojas(hojaNombre) = "NO COPIADA (no existe hoja en ORIGEN)"
            GoTo SiguienteHoja
        End If
        If wsDestino Is Nothing Then
            resultadoHojas(hojaNombre) = "NO COPIADA (no existe hoja en DESTINO)"
            GoTo SiguienteHoja
        End If

        '----------------------- Validaci�n de datos en origen --------------------------------
        ' Tomamos la �ltima fila con datos en columna A (encabezados en fila 1, datos desde fila 2)
        lastRowOrigen = wsOrigen.Cells(wsOrigen.Rows.Count, 1).End(xlUp).Row
        If lastRowOrigen < 2 Then
            resultadoHojas(hojaNombre) = "NO COPIADA (no hay datos en columna A)"
            GoTo SiguienteHoja
        End If

        '----------------------- Mapeo de encabezados (fila 1) --------------------------------
        ' Crear diccionarios: clave = texto de encabezado; valor = �ndice de columna
        Set encabezadosOrigen = CreateObject("Scripting.Dictionary")
        Set encabezadosDestino = CreateObject("Scripting.Dictionary")
        Set colsToCopy = New Collection
        columnasCopiadas = 0
        colH_Origen = 0: colH_Destino = 0

        ' Origen: localizar �ltima columna y registrar encabezados no vac�os
        lastColOrigen = wsOrigen.Cells(1, wsOrigen.Columns.Count).End(xlToLeft).Column
        For iColO = 1 To lastColOrigen
            If Trim(wsOrigen.Cells(1, iColO).Value) <> "" Then
                encabezadosOrigen(Trim(wsOrigen.Cells(1, iColO).Value)) = iColO
            End If
        Next iColO

        ' Destino: registrar encabezados no vac�os
        For iColD = 1 To wsDestino.Cells(1, wsDestino.Columns.Count).End(xlToLeft).Column
            If Trim(wsDestino.Cells(1, iColD).Value) <> "" Then
                encabezadosDestino(Trim(wsDestino.Cells(1, iColD).Value)) = iColD
            End If
        Next iColD

        '----------------------- Flags de flujo especial --------------------------------------
        esSoto = (hojaNombre = "RESIDENCIA SOTO")

        ' Si es SOTO, preparamos el mapeo especial: ORIGEN "COMISI�N" -> DESTINO "TURNO" (o "COMISI�N" si no hubiera "TURNO")
        If esSoto Then
            If encabezadosOrigen.Exists("COMISI�N") Then
                colH_Origen = encabezadosOrigen("COMISI�N")
                If encabezadosDestino.Exists("TURNO") Then
                    colH_Destino = encabezadosDestino("TURNO")
                ElseIf encabezadosDestino.Exists("COMISI�N") Then
                    colH_Destino = encabezadosDestino("COMISI�N")
                Else
                    colH_Destino = 0 ' No hay d�nde escribir; se notificar� en la validaci�n posterior
                End If
            End If
        End If

        '----------------------- Selecci�n de columnas est�ndar a copiar ----------------------
        ' Criterio: encabezado coincidente exacto y SIN f�rmulas en fila 2 (origen y destino)
        ' Nota: si es SOTO y tenemos la columna especial de "COMISI�N", la excluimos del flujo general
        For Each enc In encabezadosOrigen.Keys
            If encabezadosDestino.Exists(enc) Then
                iColO = encabezadosOrigen(enc)
                iColD = encabezadosDestino(enc)

                ' Evitar meter la columna especial en el flujo general
                If Not (esSoto And colH_Origen > 0 And iColO = colH_Origen) Then
                    tieneFormulaOrigen = wsOrigen.Cells(2, iColO).HasFormula
                    tieneFormulaDestino = wsDestino.Cells(2, iColD).HasFormula
                    If Not tieneFormulaOrigen And Not tieneFormulaDestino Then
                        colsToCopy.Add Array(iColO, iColD)
                    End If
                End If
            End If
        Next enc

        '----------------------- Validaciones previas a copiar --------------------------------
        ' Si no hay columnas est�ndar ni columna especial (cuando aplica), no hay nada que hacer
        If colsToCopy.Count = 0 And Not (esSoto And colH_Origen > 0 And colH_Destino > 0) Then
            resultadoHojas(hojaNombre) = "NO COPIADA (ninguna columna v�lida para copiar)"
            GoTo SiguienteHoja
        End If

                '======================= DETERMINAR RANGO DE FILAS A PROCESAR ==========================
        primeraFilaRango = 2
        ultimaFilaRango = lastRowOrigen
        filasCopiadas = 0
        
        ' Variable para controlar la fila de destino (consecutiva)
        Dim filaDestino As Long
        filaDestino = 2  ' Empezamos en fila 2 del destino (fila 1 = encabezados)
        
        ' Variable para evaluar RESOLUCI�N
        Dim valorResolucion As String
        
        '======================= A) Copia de columna especial (SOTO) ==========================
        ' Se ejecuta primero para evitar que otras escrituras puedan sobreescribir este resultado.
        If esSoto And colH_Origen > 0 And colH_Destino > 0 Then
            Dim arrTemp As Variant
            Dim r As Long
            
            If usarFiltro Then
                ' Con filtro:   copiar celda por celda solo las que cumplan, de forma consecutiva
                filaDestino = 2
                For filaOrigen = primeraFilaRango To ultimaFilaRango
                    cumpleFiltro = False
                    
                    ' Evaluar filtro simple o m�ltiple
                    If usarFiltroMultiple Then
                        If usarFiltroResolucion Then
                            ' Filtro m�ltiple con RESOLUCI�N:   FECHA + RESOLUCI�N
                            valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                            valorResolucion = UCase(Trim(CStr(wsOrigen.Cells(filaOrigen, criterioCol2).Value)))
                            cumpleFiltro = False
                            cumpleFiltro2 = False
                            
                            ' Evaluar fecha
                            If Not IsEmpty(valorCelda) And IsDate(valorCelda) Then
                                If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                    cumpleFiltro = True
                                End If
                            End If
                            
                            ' Evaluar RESOLUCI�N (SI, CONCEDIDA o REEVALUADA)
                            If valorResolucion = "SI" Or valorResolucion = "CONCEDIDA" Or valorResolucion = "REEVALUADA" Then
                                cumpleFiltro2 = True
                            End If
                            
                            ' Ambas condiciones deben cumplirse
                            cumpleFiltro = cumpleFiltro And cumpleFiltro2
                        Else
                            ' Filtro m�ltiple:   ENTRADA + SALIDA
                            valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                            valorCelda2 = wsOrigen.Cells(filaOrigen, criterioCol2).Value
                            cumpleFiltro = False
                            cumpleFiltro2 = False
                            
                            If Not IsEmpty(valorCelda) And IsDate(valorCelda) Then
                                If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                    cumpleFiltro = True
                                End If
                            End If
                            
                            If Not IsEmpty(valorCelda2) And IsDate(valorCelda2) Then
                                If CDate(valorCelda2) >= valorDesde2 And CDate(valorCelda2) <= valorHasta2 Then
                                    cumpleFiltro2 = True
                                End If
                            End If
                            
                            ' Ambas condiciones deben cumplirse
                            cumpleFiltro = cumpleFiltro And cumpleFiltro2
                        End If
                    Else
                        ' Filtro simple
                        valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                        If Not IsEmpty(valorCelda) Then
                            If esFechaFiltro Then
                                If IsDate(valorCelda) Then
                                    If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                        cumpleFiltro = True
                                    End If
                                End If
                            Else
                                If IsNumeric(valorCelda) Then
                                    If CDbl(valorCelda) >= valorDesde And CDbl(valorCelda) <= valorHasta Then
                                        cumpleFiltro = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                    
                    If cumpleFiltro Then
                        Dim valorComision As Variant
                        valorComision = wsOrigen.Cells(filaOrigen, colH_Origen).Value
                        If UCase$(Trim$(CStr(valorComision))) = "NO" Then
                            wsDestino.Cells(filaDestino, colH_Destino).Value = "SIN TURNO"
                        Else
                            wsDestino.Cells(filaDestino, colH_Destino).Value = valorComision
                        End If
                        filaDestino = filaDestino + 1
                    End If
                Next filaOrigen
            Else
                ' Sin filtro:  copiar bloque completo (comportamiento original)
                arrTemp = wsOrigen.Range(wsOrigen.Cells(2, colH_Origen), wsOrigen.Cells(lastRowOrigen, colH_Origen)).Value
                For r = 1 To UBound(arrTemp, 1)
                    If UCase$(Trim$(CStr(arrTemp(r, 1)))) = "NO" Then
                        arrTemp(r, 1) = "SIN TURNO"
                    End If
                Next r
                wsDestino.Range(wsDestino.Cells(2, colH_Destino), wsDestino.Cells(lastRowOrigen, colH_Destino)).Value = arrTemp
            End If
            columnasCopiadas = columnasCopiadas + 1
        End If

        '======================= B) Copia del resto de columnas est�ndar ======================
        If usarFiltro Then
            ' Con filtro:  copiar fila por fila, solo las que cumplan el criterio
            ' Las filas se copian de forma CONSECUTIVA en destino
            filaDestino = 2  ' Reiniciar para empezar desde fila 2
            
            For filaOrigen = primeraFilaRango To ultimaFilaRango
                cumpleFiltro = False
                
                ' Evaluar filtro simple o m�ltiple
                If usarFiltroMultiple Then
                    If usarFiltroResolucion Then
                        ' Filtro m�ltiple con RESOLUCI�N:   FECHA + RESOLUCI�N
                        valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                        valorResolucion = UCase(Trim(CStr(wsOrigen.Cells(filaOrigen, criterioCol2).Value)))
                        cumpleFiltro = False
                        cumpleFiltro2 = False
                        
                        ' Evaluar fecha
                        If Not IsEmpty(valorCelda) And IsDate(valorCelda) Then
                            If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                cumpleFiltro = True
                            End If
                        End If
                        
                        ' Evaluar RESOLUCI�N (SI, CONCEDIDA o REEVALUADA)
                        If valorResolucion = "SI" Or valorResolucion = "CONCEDIDA" Or valorResolucion = "REEVALUADA" Then
                            cumpleFiltro2 = True
                        End If
                        
                        ' Ambas condiciones deben cumplirse
                        cumpleFiltro = cumpleFiltro And cumpleFiltro2
                    Else
                        ' Filtro m�ltiple:  ENTRADA + SALIDA
                        valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                        valorCelda2 = wsOrigen.Cells(filaOrigen, criterioCol2).Value
                        cumpleFiltro = False
                        cumpleFiltro2 = False
                        
                        If Not IsEmpty(valorCelda) And IsDate(valorCelda) Then
                            If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                cumpleFiltro = True
                            End If
                        End If
                        
                        If Not IsEmpty(valorCelda2) And IsDate(valorCelda2) Then
                            If CDate(valorCelda2) >= valorDesde2 And CDate(valorCelda2) <= valorHasta2 Then
                                cumpleFiltro2 = True
                            End If
                        End If
                        
                        ' Ambas condiciones deben cumplirse
                        cumpleFiltro = cumpleFiltro And cumpleFiltro2
                    End If
                Else
                    ' Filtro simple
                    valorCelda = wsOrigen.Cells(filaOrigen, criterioCol).Value
                    If Not IsEmpty(valorCelda) Then
                        If esFechaFiltro Then
                            If IsDate(valorCelda) Then
                                If CDate(valorCelda) >= valorDesde And CDate(valorCelda) <= valorHasta Then
                                    cumpleFiltro = True
                                End If
                            End If
                        Else
                            If IsNumeric(valorCelda) Then
                                If CDbl(valorCelda) >= valorDesde And CDbl(valorCelda) <= valorHasta Then
                                    cumpleFiltro = True
                                End If
                            End If
                        End If
                    End If
                End If
                
                If cumpleFiltro Then
                    ' Copiar todas las columnas mapeadas para esta fila
                    ' ORIGEN:   filaOrigen -> DESTINO:  filaDestino (consecutivo)
                    For Each colPair In colsToCopy
                        iColO = colPair(0)
                        iColD = colPair(1)
                        wsDestino.Cells(filaDestino, iColD).Value = wsOrigen.Cells(filaOrigen, iColO).Value
                    Next colPair
                    filasCopiadas = filasCopiadas + 1
                    filaDestino = filaDestino + 1  ' Avanzar a la siguiente fila en destino
                End If
            Next filaOrigen
            columnasCopiadas = colsToCopy.Count
        Else
            ' Sin filtro:   copiar columnas completas (comportamiento original optimizado)
            For Each colPair In colsToCopy
                iColO = colPair(0)
                iColD = colPair(1)
                wsDestino.Range(wsDestino.Cells(2, iColD), wsDestino.Cells(lastRowOrigen, iColD)).Value = _
                    wsOrigen.Range(wsOrigen.Cells(2, iColO), wsOrigen.Cells(lastRowOrigen, iColO)).Value
                columnasCopiadas = columnasCopiadas + 1
            Next colPair
            filasCopiadas = lastRowOrigen - 1  ' Todas las filas de datos
        End If

        ' Resultado de esta hoja
        If usarFiltro Then
            If usarFiltroMultiple Then
                If usarFiltroResolucion Then
                    resultadoHojas(hojaNombre) = "COPIADA (" & filasCopiadas & " filas con filtro FECHA+RESOLUCI�N, " & columnasCopiadas & " columnas)"
                Else
                    resultadoHojas(hojaNombre) = "COPIADA (" & filasCopiadas & " filas con filtro ENTRADA+SALIDA, " & columnasCopiadas & " columnas)"
                End If
            Else
                resultadoHojas(hojaNombre) = "COPIADA (" & filasCopiadas & " filas filtradas, " & columnasCopiadas & " columnas)"
            End If
        Else
            resultadoHojas(hojaNombre) = "COPIADA CORRECTAMENTE (" & columnasCopiadas & " columnas, " & filasCopiadas & " filas)"
        End If

SiguienteHoja:
        ' Limpieza de referencias para evitar retenci�n de objetos
        Set wsOrigen = Nothing
        Set wsDestino = Nothing
        Set encabezadosOrigen = Nothing
        Set encabezadosDestino = Nothing
        Set colsToCopy = Nothing
        colH_Origen = 0: colH_Destino = 0

    Next hojaNombre

    '--------------------------- Cierre del libro origen ------------------------------------
    On Error Resume Next
    wbOrigen.Close SaveChanges:=False
    On Error GoTo 0

    '--------------------------- Resumen del proceso ----------------------------------------
    msgResumen = "RESUMEN DE COPIA DE HOJAS:" & vbCrLf & vbCrLf
    For Each hojaNombre In hojasSeleccionadas
        msgResumen = msgResumen & hojaNombre & ": " & resultadoHojas(hojaNombre) & vbCrLf
    Next hojaNombre
    MsgBox msgResumen, vbInformation, "Resultado del proceso"

'------------------------------- Restauraci�n de Excel --------------------------------------
Finalizar:
    Application.enableEvents = True
    Application.screenUpdating = True
    Application.calculation = xlCalculationAutomatic
End Sub


