# RESIDENCIAAPP — INFORME DE EVIDENCIAS TÉCNICAS (ORDEN 1W)
## VERIFICACIÓN PUNTUAL DE TRES PROCEDIMIENTOS CRÍTICOS

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Cierre del diseño arquitectónico de numeración, facturación y archivado  
**Naturaleza:** Inspección estática y extracción de código fuente, estrictamente de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico:** Antigravity  
**Fecha de emisión:** 19 de noviembre de 2026  
**Documento de referencia:** Órdenes 1S, 1T, 1U y 1V  

---

## 1. CONTROL DE FUENTES, COMMIT Y HASHES

La presente verificación técnica se ha realizado exclusivamente sobre las fuentes oficiales vigentes del repositorio y el backend de pruebas, sin aplicar modificaciones ni operaciones de escritura sobre ningún archivo o base de datos.

### 1.1. Identificación del Commit de Trabajo
- **Rama:** `master`
- **Commit HEAD:** `9f67b2c4e1dad97e8c929163d936302b0815b665`

### 1.2. Integridad de los Archivos Inspeccionados

| Archivo | Ruta Absoluta | Tamaño (Bytes) | SHA-256 |
| :--- | :--- | :---: | :--- |
| **Libro Activo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` |
| **Módulo Database** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 298.794 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` |
| **Módulo Borrado** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoDatosAlojamiento.bas` | 11.065 | `8910cfe366c08a3464caa2328460f24926a9ec8bc6c9051c20f2fc9f34fd20a0` |
| **Backend Pruebas** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7.639.040 | `9617f99795b8266c2abb91b479bf64b9ece449ad15885ce40edbe62639dbbee9` |

---

## 2. COMPROBACIÓN A — `EjecutarUpdateCondicionalOrden`

### 2.1. Localización y Contexto de Llamada
- **Archivo:** [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2448-L2505)
- **Líneas:** 2448 a 2505 (58 líneas de código).
- **Ámbito:** `Private Function`.
- **Contexto de llamada:**
  1. Es invocado directamente por `ActualizarOrden` ([modDatabase.bas#L2575](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2575)) para la modificación de campos individuales.
  2. Es invocado por `ActualizarOrdenMultiple` ([modDatabase.bas#L2636](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2636)), que a su vez es llamado por `SincronizarCambioTargetHaciaAccess` ([modDatabase.bas#L2429](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2429)) durante la edición directa en celdas de Excel interceptada por `Worksheet_Change`.

### 2.2. Código Fuente Completo Verbatim

```vba
''' Ejecuta un UPDATE condicional atómico verificando que recordsAffected > 0 (HALLAZGO-02).
''' Ejecuta un UPDATE condicional atómico exigiendo Residencia y verificando Lease Lock (HALLAZGO-02, CIERRE V5.2).
Private Function EjecutarUpdateCondicionalOrden( _
    ByVal resTarget As String, _
    ByVal numOrden As Long, _
    ByVal setParts As String, _
    Optional ByRef outUsrBloqueo As String = "", _
    Optional ByRef outSegRestantes As Long = 0, _
    Optional ByRef outExisteExpediente As Boolean = True _
) As Boolean
    Dim cn As Object
    Dim affected As Long
    Dim tokenMod As String
    Dim sql As String
    Dim whereClause As String

    On Error GoTo ErrUpdCond
    EjecutarUpdateCondicionalOrden = False
    outUsrBloqueo = ""
    outSegRestantes = 0
    outExisteExpediente = True

    If numOrden <= 0 Then Exit Function
    If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then Exit Function
    If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Function

    Call AsegurarEsquemaConcurrenciaBD

    If EstaExpedienteBloqueadoPorOtro(resTarget, numOrden, outUsrBloqueo, outSegRestantes, outExisteExpediente) Then
        Exit Function
    End If
    If Not outExisteExpediente Then Exit Function

    tokenMod = ObtenerTokenSesionInstancia()

    whereClause = " WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "' " & _
                  "AND ([UsuarioModificacion] IS NULL OR [UsuarioModificacion] = '' " & _
                  "OR [UsuarioModificacion] = '" & EscaparSQL(tokenMod) & "' " & _
                  "OR [UltimaModificacion] IS NULL " & _
                  "OR DateDiff('s', [UltimaModificacion], Now()) >= " & LEASE_LOCK_SEGUNDOS & ")"

    sql = "UPDATE Ordenes SET " & setParts & ", [UsuarioModificacion] = '" & EscaparSQL(tokenMod) & "', [UltimaModificacion] = " & FormatearFechaHoraSQL(Now) & whereClause

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()
    cn.Execute sql, affected
    cn.Close
    Set cn = Nothing

    EjecutarUpdateCondicionalOrden = (affected > 0)
    Exit Function

ErrUpdCond:
    On Error Resume Next
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    Set cn = Nothing
    On Error GoTo 0
    EjecutarUpdateCondicionalOrden = False
End Function
```

### 2.3. Análisis Técnico Detallado

#### 1. Sentencia SQL, Cláusula WHERE y Condición de Lease Lock
La sentencia ejecutada contra Access (`cn.Execute sql, affected`) se construye de forma dinámica:
```sql
UPDATE Ordenes 
SET <setParts>, 
    [UsuarioModificacion] = '<tokenMod>', 
    [UltimaModificacion] = #YYYY-MM-DD HH:NN:SS#
WHERE NumOrden = <numOrden> 
  AND Residencia = '<resTarget>'
  AND (
       [UsuarioModificacion] IS NULL 
       OR [UsuarioModificacion] = ''
       OR [UsuarioModificacion] = '<tokenMod>'
       OR [UltimaModificacion] IS NULL
       OR DateDiff('s', [UltimaModificacion], Now()) >= 120
      )
```
- **Criterio de selección:** La orden se busca combinando exclusivamente `NumOrden = <numOrden>` y `Residencia = '<resTarget>'`. **No interviene ningún discriminador de año ni clave primaria GUID**.
- **Condición de Lease Lock:** Exige que el registro no esté asignado a otro usuario activo, o pertenezca a la misma sesión (`tokenMod = ObtenerTokenSesionInstancia()`), o que hayan transcurrido 120 segundos o más (`LEASE_LOCK_SEGUNDOS = 120`) desde la última modificación.

#### 2. Mecanismo de Verificación de Filas Afectadas
- Utiliza explícitamente el parámetro de salida `affected` del método `cn.Execute sql, affected`:
  ```vba
  cn.Execute sql, affected
  ...
  EjecutarUpdateCondicionalOrden = (affected > 0)
  ```
- No se realiza una segunda consulta de verificación; se basa exclusivamente en las filas efectivamente modificadas reportadas por el motor ADO/ACE OLEDB.

#### 3. Condiciones Exactas de Devolución de True o False
- **Devuelve `False` en cualquiera de los siguientes casos:**
  1. `numOrden <= 0`.
  2. `resTarget` no es exactamente `"GIJON"`, `"SOTO"` u `"OVIEDO"`.
  3. La red/archivo Access no está accesible (`ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) = False`).
  4. `EstaExpedienteBloqueadoPorOtro` devuelve `True` (otro usuario posee el bloqueo activo dentro de la ventana de 120 segundos).
  5. `EstaExpedienteBloqueadoPorOtro` informa `outExisteExpediente = False` (el registro no existe en la tabla `Ordenes`).
  6. `affected = 0` (la condición WHERE atómica no coincidió en el momento exacto de la ejecución porque otro usuario tomó el bloqueo milisegundos antes o se eliminó la fila).
  7. Excepción en tiempo de ejecución (salto a `ErrUpdCond`).
- **Devuelve `True` exclusivamente si:**
  - El expediente existe, las validaciones iniciales son conformes, el bloqueo por Lease Lock es superado y `affected >= 1`.

#### 4. Propagación del Resultado hacia `ActualizarOrden` y el Flujo Transaccional

El procedimiento [ActualizarOrden](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2508-L2587) evalúa el resultado de `EjecutarUpdateCondicionalOrden` de la siguiente forma:

```vba
Dim usrBloq As String, segRest As Long, existeExp As Boolean
If EjecutarUpdateCondicionalOrden(resTarget, numOrden, setPart, usrBloq, segRest, existeExp) Then
    ActualizarOrden = True
Else
    If Not existeExp Then
        MsgBox "El expediente Nº ORDEN " & numOrden & " ya no existe en la base de datos.", _
               vbExclamation, "Expediente No Encontrado - " & resTarget
    ElseIf Len(usrBloq) > 0 Then
        MsgBox "El expediente Nº ORDEN " & numOrden & " está siendo editado por '" & usrBloq & "' (" & segRest & " s restantes).", _
               vbExclamation, "Bloqueo Multiusuario - " & resTarget
    End If
    ActualizarOrden = False
End If
```

- **Propagación directa al usuario:** Si la actualización fracasa, `ActualizarOrden` presenta un cuadro de mensaje modal explicativo (`MsgBox`) indicando si el expediente no existe o el nombre del usuario que lo tiene bloqueado con los segundos restantes.
- **Propagación en la edición de celdas (`Worksheet_Change` -> `SincronizarCambioTargetHaciaAccess`):**
  En `SincronizarCambioTargetHaciaAccess` ([modDatabase.bas#L2429-L2434](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2429-L2434)), si `ActualizarOrdenMultiple` (que invoca internamente a `EjecutarUpdateCondicionalOrden`) devuelve `False`:
  ```vba
  If Not ActualizarOrdenMultiple(numOrden, dictCampos, resTarget) Then
      Call SincronizarHojaDesdeAccess(resTarget, ws)
      SincronizarCambioTargetHaciaAccess = False
      Exit Function
  End If
  ```
  **Efecto crítico:** Si el update condicional fracasa por colisión o Lease Lock durante una edición directa en la cuadrícula de Excel, el sistema llama de inmediato a `SincronizarHojaDesdeAccess(resTarget, ws)`, sobreescribiendo la hoja con el estado de Access y revirtiendo la edición local del usuario.
- **Propagación en modo Contingencia Offline (SICO-N2):**
  Si la sonda de red detecta desconexión antes de llamar a `EjecutarUpdateCondicionalOrden`, `ActualizarOrden` no ejecuta el SQL, encola la instrucción en la cola offline local (`ModuloSeguridadLocal.EncolarOperacionOffline`) y devuelve `False` para evitar fingir un éxito no confirmado por el backend.

---

## 3. COMPROBACIÓN B — `LimpiarTablaAlojamientoCompleta`
*(Procedimiento real: `BorrarDatosAlojamientoOptimizado`)*

### 3.1. Identificación y Nombre Exacto del Procedimiento
- **Archivo:** [BorradoDatosAlojamiento.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas#L21-L244)
- **Líneas:** 21 a 244 (224 líneas de código).
- **Nombre real en código:** `Public Sub BorrarDatosAlojamientoOptimizado(control As IRibbonControl)`.
- **Aclaración sobre el nombre documental:** En diversos análisis y pliegos arquitectónicos previos se le referenció funcionalmente como `LimpiarTablaAlojamientoCompleta`. No existe en el proyecto ningún procedimiento con ese nombre literal; el procedimiento unificado y oficial que ejecuta el vaciado masivo de las tablas de alojamiento en las hojas activas es `BorrarDatosAlojamientoOptimizado`.

### 3.2. Código Fuente Completo Verbatim

```vba
Public Sub BorrarDatosAlojamientoOptimizado(control As IRibbonControl)
    ' Validar permisos de Administrador (Fase 1 - Paso 7)
    If Not modDatabase.ValidarPermisoAdmin("Borrado Masivo de Alojamientos") Then Exit Sub

    On Error GoTo ErrHandler

    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' --- Confirmaciones ---
    If MsgBox("Se va a proceder al borrado de la totalidad de los datos de la hoja '" & ws.Name & "'." & vbCrLf & vbCrLf & _
              "¿Desea continuar?", vbExclamation + vbYesNo, "Confirmación Inicial") <> vbYes Then Exit Sub

    If MsgBox("Realmente desea borrar la totalidad de los datos de la hoja '" & ws.Name & "'?" & vbCrLf & vbCrLf & _
              "El proceso es IRREVERSIBLE." & vbCrLf & vbCrLf & "¿Confirma el borrado completo?", _
              vbCritical + vbYesNo, "Confirmación Final - ¡ATENCIÓN!") <> vbYes Then Exit Sub

    ' --- Columnas protegidas por hoja ---
    Dim protectedCols As Variant
    Dim nomHojaCal As String, nomHojaSum As String
    Select Case ws.Name
        Case "RESIDENCIA GIJÓN"
            protectedCols = Array("U", "AB", "AC")
            nomHojaCal = "Calendario GIJÓN"
            nomHojaSum = "RESUMEN GIJÓN"
        Case "RESIDENCIA SOTO"
            protectedCols = Array("Q", "R", "U", "AB", "AC")
            nomHojaCal = "Calendario SOTO"
            nomHojaSum = "RESUMEN SOTO"
        Case "RESIDENCIA OVIEDO"
            protectedCols = Array("V", "AF", "AG")
            nomHojaCal = "Calendario OVIEDO"
            nomHojaSum = "RESUMEN OVIEDO"
        Case Else
            MsgBox "Esta macro solo puede ejecutarse en las hojas de alojamiento", vbCritical, "Error"
            Exit Sub
    End Select

    ' --- Optimización de Excel ---
    Application.screenUpdating = False
    Application.enableEvents = False
    Application.calculation = xlCalculationManual

    ' --- Aislar recálculo de hojas satélite pesadas (Calendario y RESUMEN) para evitar "Excel no responde" ---
    Dim wsCal As Worksheet, wsSum As Worksheet
    Set wsCal = modDatabase.ObtenerHojaSegura(nomHojaCal)
    Set wsSum = modDatabase.ObtenerHojaSegura(nomHojaSum)

    ' --- Desproteger hoja y quitar filtros ANTES de medir lastRow ---
    Dim estabaProtegida As Boolean
    Dim pwdHojas As String
    estabaProtegida = ws.ProtectContents
    pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
    If estabaProtegida Then
        On Error Resume Next
        If Len(pwdHojas) > 0 Then ws.Unprotect password:=pwdHojas
        If ws.ProtectContents Then ws.Unprotect password:=""
        If ws.ProtectContents Then ws.Unprotect
        On Error GoTo ErrHandler
    End If

    If ws.FilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo ErrHandler
    End If

    ' --- Cálculo de última fila real examinando SOLO columnas de datos (excluyendo fórmulas pre-arrastradas) ---
    Dim lastCol As Long
    Select Case ws.Name
        Case "RESIDENCIA OVIEDO": lastCol = 31 ' Datos hasta AE (excluye V=22, AF=32, AG=33)
        Case Else:                lastCol = 30 ' Datos hasta AD (excluye Q=17, R=18 en Soto y U=21, AB=28, AC=29)
    End Select

    Dim lastRow As Long, i As Long, tempRow As Long
    Dim esColFormula As Boolean
    lastRow = 2
    For i = 1 To lastCol
        esColFormula = False
        Select Case ws.Name
            Case "RESIDENCIA GIJÓN"
                If i = 21 Or i = 28 Or i = 29 Then esColFormula = True
            Case "RESIDENCIA SOTO"
                If i = 17 Or i = 18 Or i = 21 Or i = 28 Or i = 29 Then esColFormula = True
            Case "RESIDENCIA OVIEDO"
                If i = 22 Then esColFormula = True
        End Select

        If Not esColFormula Then
            tempRow = ws.Cells(ws.Rows.Count, i).End(xlUp).Row
            If tempRow > lastRow Then lastRow = tempRow
        End If
    Next i

    ' --- Saneamiento ultraligero de columnas excedentes (> AJ=36) sin desplazar 16.348 columnas enteras ---
    On Error Resume Next
    Dim ultColUsada As Long, ultFilaUsada As Long
    ultColUsada = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultColUsada > 36 Then
        ultFilaUsada = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If ultFilaUsada < lastRow Then ultFilaUsada = lastRow
        ws.Range(ws.Cells(1, 37), ws.Cells(ultFilaUsada, ultColUsada)).Clear
    End If
    On Error GoTo ErrHandler

    ' --- Recolectar órdenes activas en RAM para soporte SICO-N2 Offline ---
    Dim colOrdenesHoja As New Collection
    If lastRow >= 2 Then
        Dim arrCola As Variant, rOrd As Long
        If lastRow = 2 Then
            ReDim arrCola(1 To 1, 1 To 1)
            arrCola(1, 1) = ws.Cells(2, 1).Value
        Else
            arrCola = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 1)).Value
        End If
        For rOrd = 1 To UBound(arrCola, 1)
            If IsNumeric(arrCola(rOrd, 1)) And Len(Trim$(CStr(arrCola(rOrd, 1)))) > 0 Then
                If CLng(arrCola(rOrd, 1)) > 0 Then colOrdenesHoja.Add CLng(arrCola(rOrd, 1))
            End If
        Next rOrd
    End If

    ' --- SINCRONIZACIÓN ACCESS / CONTINGENCIA SICO-N2 EN 1 SOLA TRANSACCIÓN MASIVA ---
    Dim claveResTotal As String
    Dim totalBorradasBD As Long
    claveResTotal = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
    totalBorradasBD = 0

    If Len(claveResTotal) > 0 Then
        If Not modDatabase.EliminarTodasOrdenesSedeBD(claveResTotal, totalBorradasBD, colOrdenesHoja) Then
            If estabaProtegida Then
                On Error Resume Next
                ws.Protect password:=pwdHojas, UserInterfaceOnly:=True, _
                    AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
                    AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
                On Error GoTo ErrHandler
            End If
            Application.enableEvents = True
            Application.calculation = xlCalculationAutomatic
            Application.DisplayAlerts = True
            Application.screenUpdating = True
            Exit Sub
        End If
        Call modDatabase.InsertarLog(modDatabase.ObtenerUsuarioActual(), "BORRADO_TOTAL_TABLA", _
            "Hoja: " & ws.Name & " | Órdenes eliminadas/encoladas: " & totalBorradasBD)
    End If

    ' --- DisplayAlerts JUSTO antes de borrar (evita avisos de celdas bloqueadas) ---
    Application.DisplayAlerts = False

    ' --- Pausar evaluación de hojas satélite (Calendario y RESUMEN) durante el vaciado ---
    On Error Resume Next
    If Not wsCal Is Nothing Then wsCal.EnableCalculation = False
    If Not wsSum Is Nothing Then wsSum.EnableCalculation = False
    On Error GoTo ErrHandler

    ' --- Borrado masivo por bloques contiguos respetando columnas protegidas ---
    Select Case ws.Name
        Case "RESIDENCIA GIJÓN"
            ' Protegidas: U (21), AB (28), AC (29). Limpieza de 1..20, 22..27 y 30 (Repartir Coste)
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 20)).ClearContents
            ws.Range(ws.Cells(2, 22), ws.Cells(lastRow, 27)).ClearContents
            ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents
        Case "RESIDENCIA SOTO"
            ' Protegidas: Q (17), R (18), U (21), AB (28), AC (29). Limpieza de 1..16, 19..20, 22..27 y 30
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 16)).ClearContents
            ws.Range(ws.Cells(2, 19), ws.Cells(lastRow, 20)).ClearContents
            ws.Range(ws.Cells(2, 22), ws.Cells(lastRow, 27)).ClearContents
            ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents
        Case "RESIDENCIA OVIEDO"
            ' Protegidas: V (22), AF (32), AG (33). Limpieza de 1..21 y 23..31
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 21)).ClearContents
            ws.Range(ws.Cells(2, 23), ws.Cells(lastRow, 31)).ClearContents
    End Select

    ' --- Limpieza de comentarios columna P ---
    If lastRow >= 2 Then
        ws.Range("P2:P" & lastRow).ClearComments
    End If

    ' --- Restaurar protección de la hoja si estaba protegida ---
    If estabaProtegida Then
        On Error Resume Next
        ws.Protect password:=pwdHojas, UserInterfaceOnly:=True, _
            AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
            AllowInsertingRows:=True, AllowDeletingRows:=True, AllowSorting:=True, AllowFiltering:=True
        On Error GoTo ErrHandler
    End If

    ' --- Recálculo escalonado cediendo mensajes a Windows (evita "Excel no responde" en Gijón y Soto) ---
    On Error Resume Next
    ws.Calculate
    DoEvents
    If Not wsSum Is Nothing Then wsSum.EnableCalculation = True
    DoEvents
    If Not wsCal Is Nothing Then wsCal.EnableCalculation = True
    On Error GoTo ErrHandler

    ' --- Restauración del entorno Excel ---
    Application.calculation = xlCalculationAutomatic
    Application.enableEvents = True
    Application.DisplayAlerts = True
    Application.screenUpdating = True
    DoEvents

    ' --- Mensaje de resumen ---
    MsgBox "Filas procesadas: " & (lastRow - 1) & vbCrLf & _
           "Estado protección: " & IIf(ws.ProtectContents, "PROTEGIDA", "DESPROTEGIDA"), _
           vbInformation, "Limpieza Finalizada - " & ws.Name
    Exit Sub

ErrHandler:
    ' --- Restauración segura ante error ---
    On Error Resume Next
    If Not wsSum Is Nothing Then wsSum.EnableCalculation = True
    If Not wsCal Is Nothing Then wsCal.EnableCalculation = True
    On Error GoTo 0
    Application.DisplayAlerts = True
    Application.enableEvents = True
    Application.calculation = xlCalculationAutomatic
    Application.screenUpdating = True

    MsgBox "Error durante el borrado:" & vbCrLf & Err.Description, vbCritical, "Error en " & ws.Name
End Sub
```

### 3.3. Columnas y Rangos Limpiados por Residencia

| Residencia | Rangos de Filas | Columnas Limpiadas (`ClearContents`) | Columnas de Fórmulas Blindadas (Intactas) |
| :--- | :---: | :--- | :--- |
| **GIJÓN** | `2` a `lastRow` | **Bloque 1:** Cols 1 a 20 (`A:T`)<br>**Bloque 2:** Cols 22 a 27 (`V:AA`)<br>**Bloque 3:** Col 30 (`AD` = `REPARTIR COSTE  HAB.`) | **Col 21 (`U`):** `IMPORTE`<br>**Col 28 (`AB`):** `NOMBRE PDF`<br>**Col 29 (`AC`):** `TEXTO EMAIL` |
| **SOTO** | `2` a `lastRow` | **Bloque 1:** Cols 1 a 16 (`A:P`)<br>**Bloque 2:** Cols 19 a 20 (`S:T`)<br>**Bloque 3:** Cols 22 a 27 (`V:AA`)<br>**Bloque 4:** Col 30 (`AD` = `REPARTIR COSTE  APTO.`) | **Col 17 (`Q`):** `PRECIO APTO`<br>**Col 18 (`R`):** `SUPLE OCUPAN`<br>**Col 21 (`U`):** `IMPORTE`<br>**Col 28 (`AB`):** `NOMBRE PDF`<br>**Col 29 (`AC`):** `TEXTO EMAIL` |
| **OVIEDO** | `2` a `lastRow` | **Bloque 1:** Cols 1 a 21 (`A:U`)<br>**Bloque 2:** Cols 23 a 31 (`W:AE`) *(incluye `AC`=29, `AD`=30 y `AE`=31)* | **Col 22 (`V`):** `IMPORTE`<br>**Col 32 (`AF`):** `NOMBRE PDF`<br>**Col 33 (`AG`):** `TEXTO EMAIL` |

### 3.4. Resolución de la Discrepancia entre las Secciones 3.3 y 8 del Informe 1V

- **Origen de la duda en el Informe 1V:**
  - En la **Sección 3.3 del Informe 1V** se determinó que los 5 campos de usuario locales (`REPARTIR COSTE` en GIJÓN, SOTO y OVIEDO, junto con `TAG` y `R` en OVIEDO) **no tienen persistencia en Access** (no existen columnas para ellos en la tabla `Ordenes`).
  - En la **Sección 8 del Informe 1V**, al analizar la persistencia local frente al vaciado y la recarga, surgió una aparente contradicción entre si estos campos eran borrados o no al vaciar la tabla.
- **Evidencia Técnica Concluyente:**
  1. En este procedimiento de vaciado masivo ([BorradoDatosAlojamiento.bas#L178-L195](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas#L178-L195)), **LOS CINCO CAMPOS DE USUARIO SÍ SE BORRAN EXPRESAMENTE**:
     - GIJÓN: Columna 30 (`AD`, `REPARTIR COSTE  HAB.`) se borra explícitamente mediante `ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents`.
     - SOTO: Columna 30 (`AD`, `REPARTIR COSTE  APTO.`) se borra explícitamente mediante `ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents`.
     - OVIEDO: El rango `ws.Range(ws.Cells(2, 23), ws.Cells(lastRow, 31)).ClearContents` cubre desde la columna 23 (`W`) hasta la 31 (`AE`), limpiando en un solo bloque continuo las columnas 29 (`AC`, `REPARTIR COSTE  HAB.`), 30 (`AD`, `TAG`) y 31 (`AE`, `R`).
  2. **¿Por qué existía la discrepancia?**
     Porque se estaba comparando inadvertidamente el comportamiento de este vaciado masivo con el procedimiento de sincronización y descarga desde Access (`SincronizarHojaDesdeAccess`, analizado en la Comprobación C).
     En `SincronizarHojaDesdeAccess`, la limpieza previa **NO limpia las columnas 30 de GIJÓN/SOTO ni las 29..31 de OVIEDO**, dejándolas intactas mientras sobreescribe las demás.
  3. **Conclusión:** En `BorrarDatosAlojamientoOptimizado` **ninguna columna de datos de usuario queda sin limpiar**. Las únicas columnas que se preservan intactas son única y exclusivamente las que contienen fórmulas calculadas (`U, AB, AC` en GIJÓN; `Q, R, U, AB, AC` en SOTO; `V, AF, AG` en OVIEDO). Tampoco se eliminan físicamente las filas (no se usa `Delete Shift:=xlUp`), sino `ClearContents`, manteniendo intacta la geometría de la cuadrícula.

---

## 4. COMPROBACIÓN C — DESCARGA Y RECONSTRUCCIÓN DESDE ACCESS
*(Procedimiento real: `SincronizarHojaDesdeAccess`)*

### 4.1. Localización y Nombre Exacto del Procedimiento
- **Archivo:** [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3235-L3602)
- **Líneas:** 3235 a 3602 (368 líneas de código).
- **Nombre real en código:** `Public Sub SincronizarHojaDesdeAccess(ByVal residencia As String, ByVal hoja As Worksheet, Optional ByVal filaInicio As Long = 2)`.
- **Aclaración sobre `DescargarOrdenesDesdeAccess`:** Dicho nombre es una designación documental histórica. En todo el repositorio no existe ningún procedimiento denominado `DescargarOrdenesDesdeAccess`; el procedimiento real y único que descarga y reconstruye los expedientes de una sede desde la tabla `Ordenes` hacia las hojas de Excel es `SincronizarHojaDesdeAccess`.

### 4.2. Catálogo Completo de Llamadas en el Proyecto
El procedimiento es el pilar central de refresco y reconstrucción de la aplicación, ejecutado en 22 puntos estratégicos:
1. **Activación de Hojas de Residencia:**
   - [Hoja2.cls#L56](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja2.cls#L56) (`Worksheet_Activate` de RESIDENCIA GIJÓN).
   - [Hoja20.cls#L59](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja20.cls#L59) (`Worksheet_Activate` de RESIDENCIA SOTO).
   - [Hoja24.cls#L31](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja24.cls#L31) (`Worksheet_Activate` de RESIDENCIA OVIEDO).
2. **Reversión tras Error en Edición Directa:**
   - [modDatabase.bas#L2364](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2364), [L2372](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2372), [L2430](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2430) en `SincronizarCambioTargetHaciaAccess`.
3. **Confirmación de Altas y Reservas Transaccionales:**
   - [modDatabase.bas#L1631](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1631) en `ConfirmarReservaConTransaccion`.
   - [modDatabase.bas#L1839](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1839), [L1856](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1856) en `InsertarNuevaOrdenConTransaccion`.
4. **Eliminación y Archivado de Expedientes:**
   - [modDatabase.bas#L4349](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4349) tras borrar una orden en `EliminarOrdenBD`.
5. **Cinta de Opciones (Ribbon) y Botones de Actualización:**
   - [modDatabase.bas#L3630](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3630), [L3649](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3649) en `SincronizarResidenciaActiva`.
   - [modDatabase.bas#L3686](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3686) en `ActualizarDesdeBDAS_HojasResidencias_Ribbon`.
   - [BotonIrUltimaCeldaOrden.bas#L17](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonIrUltimaCeldaOrden.bas#L17).
   - [BusquedaOrdenNombreFactura.bas#L118](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BusquedaOrdenNombreFactura.bas#L118).
   - [DuplicadosOrdenyFactura.bas#L371](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/DuplicadosOrdenyFactura.bas#L371).
6. **Módulos de Facturación y Calendarios:**
   - [FacturacionMesGIJON.bas#L43](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/FacturacionMesGIJON.bas#L43), [FacturacionMesOVIEDO.bas#L32](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/FacturacionMesOVIEDO.bas#L32), [FacturacionMesSOTO.bas#L32](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/FacturacionMesSOTO.bas#L32).
   - [ModuloCalendarioGijon.bas#L74](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloCalendarioGijon.bas#L74), [ModuloCalendarioOviedo.bas#L131](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloCalendarioOviedo.bas#L131), [ModuloCalendarioSoto.bas#L84](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloCalendarioSoto.bas#L84).
   - [ValidacionFechasGijon.bas#L123](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ValidacionFechasGijon.bas#L123), [ValidacionFechasOviedo.bas#L57](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ValidacionFechasOviedo.bas#L57), [ValidacionFechasSoto.bas#L79](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ValidacionFechasSoto.bas#L79).

### 4.3. Código Fuente Completo Verbatim

```vba
Public Sub SincronizarHojaDesdeAccess( _
    ByVal residencia As String, _
    ByVal hoja As Worksheet, _
    Optional ByVal filaInicio As Long = 2 _
)
    Dim rs As Object
    Dim sql As String
    Dim ultimaFila As Long
    
    On Error GoTo ErrorHandler
    
    ' === [SICO-N2: BARRERA DE INTEGRIDAD ANTES DE SOBREESCRIBIR EXCEL DESDE ACCESS] ===
    If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
        Exit Sub
    End If
    If ModuloSeguridadLocal.ContarCambiosPendientes() > 0 And Not ModuloSeguridadLocal.EstaSincronizandoFIFO() Then
        Call ModuloSeguridadLocal.SincronizarColaOfflineHaciaAccess
    End If
    
    Dim scrOrig As Boolean, evOrig As Boolean
    scrOrig = Application.screenUpdating
    evOrig = Application.enableEvents

    Application.screenUpdating = False
    Application.enableEvents = False  ' Evitar que Worksheet_Change se dispare durante la carga
    
    Dim estabaProtegida As Boolean
    estabaProtegida = hoja.ProtectContents
    If estabaProtegida Then
        On Error Resume Next
        Dim pwdHojas As String
        pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
        If Len(pwdHojas) > 0 Then
            hoja.Unprotect password:=pwdHojas
        End If
        If hoja.ProtectContents Then
            hoja.Unprotect password:=""
        End If
        If hoja.ProtectContents Then
            hoja.Unprotect
        End If
        On Error GoTo ErrorHandler
    End If

    sql = "SELECT * FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "' ORDER BY Id ASC"
    Set rs = GetRecordset(sql)
    
    If rs Is Nothing Then
        GoTo CleanUp
    End If
    
    ' Si hay un filtro activo en la hoja, mostrar todas las filas antes de limpiar y volcar en bloque
    On Error Resume Next
    If hoja.FilterMode Then hoja.ShowAllData
    On Error GoTo ErrorHandler
    
    Dim resNorm As String
    resNorm = UCase(Trim(residencia))
    
    Dim isOviedo As Boolean, isSoto As Boolean, isGijon As Boolean
    isOviedo = (resNorm = "OVIEDO")
    isSoto = (resNorm = "SOTO")
    isGijon = (resNorm = "GIJON" Or resNorm = "GIJÓN")
    
    ' Limpiar datos antiguos de la hoja preservando estrictamente las columnas de fórmulas
    ultimaFila = hoja.Cells(hoja.Rows.Count, 1).End(xlUp).Row
    If ultimaFila >= filaInicio Then
        If isOviedo Then
            ' OVIEDO: Limpiar solo columnas de datos A:U (1:21) y W:AB (23:28).
            ' Fórmulas blindadas: V (22, IMPORTE), AF (32, NOMBRE PDF), AG (33, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 21)).ClearContents    ' Cols A a U
            hoja.Range(hoja.Cells(filaInicio, 23), hoja.Cells(ultimaFila, 28)).ClearContents   ' Cols W a AB
        ElseIf isSoto Then
            ' SOTO: Limpiar solo columnas de datos A:P (1:16), S:T (19:20) y V:AA (22:27).
            ' Fórmulas blindadas: Q (17, PRECIO APTO), R (18, SUPLE OCUPAN), U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 16)).ClearContents    ' Cols A a P
            hoja.Range(hoja.Cells(filaInicio, 19), hoja.Cells(ultimaFila, 20)).ClearContents   ' Cols S a T
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(ultimaFila, 27)).ClearContents   ' Cols V a AA
        Else
            ' GIJÓN: Limpiar solo columnas de datos A:T (1:20) y V:AA (22:27).
            ' Fórmulas blindadas: U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL).
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(ultimaFila, 20)).ClearContents    ' Cols A a T
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(ultimaFila, 27)).ClearContents   ' Cols V a AA
        End If
    End If
    
    If rs.EOF Then
        rs.Close
        Dim errBloqEOF As String
        If Not CargarBloqueosYEstadoHabitacionesDesdeBD(resNorm, errBloqEOF) Then
            MsgBox "Se sincronizaron las órdenes de " & resNorm & ", pero no se pudieron recargar los bloqueos/habitaciones:" & _
                   vbCrLf & vbCrLf & "  • Detalle: " & errBloqEOF, vbExclamation, "Aviso de Sincronización - " & resNorm
        End If
        GoTo CleanUp
    End If
    
    Dim vVal As Variant
    
    ' Obtener total de registros
    Dim totalRows As Long
    rs.MoveLast
    totalRows = rs.RecordCount
    rs.MoveFirst
    
    If totalRows > 0 Then
        Dim r As Long
        
        If isOviedo Then
            ' OVIEDO: Bloque 1 (Cols A a U = 1 a 21), Bloque 2 (Cols W a AB = 23 a 28)
            ' Col V (22, IMPORTE), Col AF (32, FÓRMULA PDF) y Col AG (33, FÓRMULA EMAIL) se mantienen intactas
            Dim arrOviedo1() As Variant
            Dim arrOviedo2() As Variant
            ReDim arrOviedo1(1 To totalRows, 1 To 21)
            ReDim arrOviedo2(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrOviedo1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrOviedo1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrOviedo1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÓM FACT)
                arrOviedo1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrOviedo1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrOviedo1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrOviedo1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Comision")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Turno")
                arrOviedo1(r, 8) = vVal                                                 ' Col H (COMISIÓN / TURNO)
                
                arrOviedo1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrOviedo1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrOviedo1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrOviedo1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrOviedo1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrOviedo1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrOviedo1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrOviedo1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                arrOviedo1(r, 17) = ObtenerValorCampo(rs, "NumHabIndividuales")         ' Col Q (HAB. IND.)
                arrOviedo1(r, 18) = ObtenerValorCampo(rs, "NumHabDobles")               ' Col R (HAB. DOBLE)
                arrOviedo1(r, 19) = ObtenerValorCampo(rs, "CamaSuple")                  ' Col S (CAMA SUPLE.)
                arrOviedo1(r, 20) = ObtenerValorCampo(rs, "HabitacionesAsignadas")      ' Col T (NÓM HAB.)
                arrOviedo1(r, 21) = ObtenerValorCampo(rs, "DtoFamNum")                  ' Col U (DTO. FAM. NUM.)
                
                ' Col V (22): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrOviedo2(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col W (TELEFONO)
                arrOviedo2(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col X (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrOviedo2(r, 3) = vVal                                                 ' Col Y (CP)
                
                arrOviedo2(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Z (POBLACIÓN)
                arrOviedo2(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col AA (PROVINCIA)
                arrOviedo2(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AB (PAGADO)
                
                ' Col AF (32): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AG (33): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 21)).Value = arrOviedo1
            hoja.Range(hoja.Cells(filaInicio, 23), hoja.Cells(filaInicio + totalRows - 1, 28)).Value = arrOviedo2
            
        ElseIf isSoto Then
            ' SOTO: Bloque 1 (Cols A a P = 1 a 16), Bloque 2 (Cols S a T = 19 a 20), Bloque 3 (Cols V a AA = 22 a 27)
            ' Col Q (17, PRECIO APTO), Col R (18, SUPLE OCUPAN), Col U (21, IMPORTE),
            ' Col AB (28, NOMBRE PDF) y Col AC (29, TEXTO EMAIL) se mantienen 100% INTACTAS con sus fórmulas
            Dim arrSoto1() As Variant
            Dim arrSoto2() As Variant
            Dim arrSoto3() As Variant
            ReDim arrSoto1(1 To totalRows, 1 To 16)
            ReDim arrSoto2(1 To totalRows, 1 To 2)
            ReDim arrSoto3(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrSoto1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrSoto1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrSoto1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÓM FACT)
                arrSoto1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrSoto1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrSoto1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrSoto1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Turno")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Comision")
                arrSoto1(r, 8) = vVal                                                 ' Col H (TURNO)
                
                arrSoto1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrSoto1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrSoto1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrSoto1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrSoto1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrSoto1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrSoto1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrSoto1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                
                ' Col Q (17): PRECIO APTO (FÓRMULA PRESERVADA INTACTA)
                ' Col R (18): SUPLE OCUPAN (FÓRMULA PRESERVADA INTACTA)
                
                arrSoto2(r, 1) = ObtenerValorCampo(rs, "HabitacionesAsignadas")       ' Col S (APARTAMENTO)
                arrSoto2(r, 2) = ObtenerValorCampo(rs, "DtoFamNum")                   ' Col T (DTO. FAM. NUM.)
                
                ' Col U (21): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrSoto3(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col V (TELEFONO)
                arrSoto3(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col W (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrSoto3(r, 3) = vVal                                                 ' Col X (CP)
                
                arrSoto3(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Y (POBLACIÓN)
                arrSoto3(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col Z (PROVINCIA)
                arrSoto3(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AA (PAGADO)
                
                ' Col AB (28): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AC (29): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 16)).Value = arrSoto1
            hoja.Range(hoja.Cells(filaInicio, 19), hoja.Cells(filaInicio + totalRows - 1, 20)).Value = arrSoto2
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(filaInicio + totalRows - 1, 27)).Value = arrSoto3
            
        Else
            ' GIJÓN: Bloque 1 (Cols A a T = 1 a 20), Bloque 2 (Cols V a AA = 22 a 27)
            ' Col U (21, IMPORTE), Col AB (28, NOMBRE PDF) y Col AC (29, TEXTO EMAIL) se mantienen 100% INTACTAS con sus fórmulas
            Dim arrGijon1() As Variant
            Dim arrGijon2() As Variant
            ReDim arrGijon1(1 To totalRows, 1 To 20)
            ReDim arrGijon2(1 To totalRows, 1 To 6)
            
            For r = 1 To totalRows
                arrGijon1(r, 1) = ObtenerValorCampo(rs, "NumOrden")                   ' Col A (Nº ORDEN)
                arrGijon1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")               ' Col B (FECHA PETICION)
                arrGijon1(r, 3) = ObtenerValorCampo(rs, "NumFactura")                  ' Col C (NÓM FACT)
                arrGijon1(r, 4) = ObtenerValorCampo(rs, "Finalidad")                   ' Col D (FINALIDAD)
                arrGijon1(r, 5) = ObtenerValorCampo(rs, "Empleo")                      ' Col E (EMPLEO)
                arrGijon1(r, 6) = ObtenerValorCampo(rs, "Situacion")                   ' Col F (SITUACION)
                arrGijon1(r, 7) = ObtenerValorCampo(rs, "Evaluacion")                  ' Col G (EVALUACIÓN)
                
                vVal = ObtenerValorCampo(rs, "Comision")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "Turno")
                arrGijon1(r, 8) = vVal                                                 ' Col H (COMISIÓN)
                
                arrGijon1(r, 9) = ObtenerValorCampo(rs, "DNI")                         ' Col I (DNI)
                arrGijon1(r, 10) = ObtenerValorCampo(rs, "Rango")                      ' Col J (RANGO)
                arrGijon1(r, 11) = ObtenerValorCampo(rs, "Nombre")                     ' Col K (Nombre)
                arrGijon1(r, 12) = ObtenerValorCampo(rs, "FechaEntrada")               ' Col L (ENTRADA)
                arrGijon1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")                ' Col M (SALIDA)
                arrGijon1(r, 14) = ObtenerValorCampo(rs, "DiasUso")                    ' Col N (DIAS USO)
                arrGijon1(r, 15) = ObtenerValorCampo(rs, "PAX")                        ' Col O (PAX)
                arrGijon1(r, 16) = ObtenerValorCampo(rs, "Resolucion")                 ' Col P (RESOLUCION)
                arrGijon1(r, 17) = ObtenerValorCampo(rs, "NumHabIndividuales")         ' Col Q (HAB. IND.)
                arrGijon1(r, 18) = ObtenerValorCampo(rs, "NumHabDobles")               ' Col R (HAB. DOBLE)
                arrGijon1(r, 19) = ObtenerValorCampo(rs, "HabitacionesAsignadas")      ' Col S (NÓM HAB.)
                arrGijon1(r, 20) = ObtenerValorCampo(rs, "DtoFamNum")                  ' Col T (DTO. FAM. NUM.)
                
                ' Col U (21): IMPORTE (FÓRMULA PRESERVADA INTACTA)
                
                arrGijon2(r, 1) = ObtenerValorCampo(rs, "Telefono")                    ' Col V (TELEFONO)
                arrGijon2(r, 2) = ObtenerValorCampo(rs, "Direccion")                   ' Col W (DIRECCIÓN)
                
                vVal = ObtenerValorCampo(rs, "CP")
                If IsNull(vVal) Then vVal = ObtenerValorCampo(rs, "CodigoPostal")
                arrGijon2(r, 3) = vVal                                                 ' Col X (CP)
                
                arrGijon2(r, 4) = ObtenerValorCampo(rs, "Poblacion")                   ' Col Y (POBLACIÓN)
                arrGijon2(r, 5) = ObtenerValorCampo(rs, "Provincia")                   ' Col Z (PROVINCIA)
                arrGijon2(r, 6) = ObtenerValorCampo(rs, "EstadoPago")                  ' Col AA (PAGADO)
                
                ' Col AB (28): FÓRMULA NOMBRE PDF (PRESERVADA INTACTA)
                ' Col AC (29): FÓRMULA COPIAR TEXTO EMAIL (PRESERVADA INTACTA)
                
                rs.MoveNext
            Next r
            
            ' Volcado ultra-rápido en bloques COM preservando columnas de fórmulas
            hoja.Range(hoja.Cells(filaInicio, 1), hoja.Cells(filaInicio + totalRows - 1, 20)).Value = arrGijon1
            hoja.Range(hoja.Cells(filaInicio, 22), hoja.Cells(filaInicio + totalRows - 1, 27)).Value = arrGijon2
            
        End If
        
        ' =========================================================================
        ' BLINDAJE Y AUTO-REPARACIÓN UNIVERSAL DE FÓRMULAS POST-SINCRONIZACIÓN
        ' Garantiza que cada fila volcada conserve o repare sus fórmulas maestras
        ' =========================================================================
        Dim colsFormula() As Long
        Dim idxF As Long, cCol As Long
        Dim filaRef As Long, fRow As Long
        Dim fmlaMaster As String
        
        If isOviedo Then
            ' OVIEDO: V (22, IMPORTE), AF (32, NOMBRE PDF), AG (33, TEXTO EMAIL)
            ReDim colsFormula(1 To 3)
            colsFormula(1) = 22
            colsFormula(2) = 32
            colsFormula(3) = 33
        ElseIf isSoto Then
            ' SOTO: Q (17, PRECIO APTO), R (18, SUPLE OCUPAN), U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL)
            ReDim colsFormula(1 To 5)
            colsFormula(1) = 17
            colsFormula(2) = 18
            colsFormula(3) = 21
            colsFormula(4) = 28
            colsFormula(5) = 29
        Else
            ' GIJÓN: U (21, IMPORTE), AB (28, NOMBRE PDF), AC (29, TEXTO EMAIL)
            ReDim colsFormula(1 To 3)
            colsFormula(1) = 21
            colsFormula(2) = 28
            colsFormula(3) = 29
        End If
        
        For idxF = LBound(colsFormula) To UBound(colsFormula)
            cCol = colsFormula(idxF)
            filaRef = 0
            
            ' 1. Probar en las primeras filas operativas
            If hoja.Cells(filaInicio, cCol).HasFormula Then
                filaRef = filaInicio
            ElseIf hoja.Cells(filaInicio + 1, cCol).HasFormula Then
                filaRef = filaInicio + 1
            Else
                ' 2. Buscar en filas de plantilla posteriores no afectadas por la carga
                Dim rScan As Long
                For rScan = filaInicio + totalRows To filaInicio + totalRows + 100
                    If hoja.Cells(rScan, cCol).HasFormula Then
                        filaRef = rScan
                        Exit For
                    End If
                Next rScan
            End If
            
            ' Si encontramos la fórmula maestra de referencia, propagarla a toda celda que carezca de ella
            If filaRef > 0 Then
                fmlaMaster = hoja.Cells(filaRef, cCol).FormulaR1C1
                hoja.Range(hoja.Cells(filaInicio, cCol), hoja.Cells(filaInicio + totalRows - 1, cCol)).FormulaR1C1 = fmlaMaster
            End If
        Next idxF
    End If
    
    rs.Close
    Dim errBloqFin As String
    If Not CargarBloqueosYEstadoHabitacionesDesdeBD(resNorm, errBloqFin) Then
        MsgBox "Se sincronizaron las órdenes de " & resNorm & ", pero no se pudieron recargar los bloqueos/habitaciones:" & _
               vbCrLf & vbCrLf & "  • Detalle: " & errBloqFin, vbExclamation, "Aviso de Sincronización - " & resNorm
    End If
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error al sincronizar datos desde Access:" & vbCrLf & _
           Err.Description, vbCritical, "Error de Sincronización"

CleanUp:
    If estabaProtegida Then
        On Error Resume Next
        hoja.Protect password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
        On Error GoTo 0
    End If
    Application.enableEvents = evOrig
    Application.screenUpdating = scrOrig
    If Not rs Is Nothing Then Set rs = Nothing
End Sub
```

### 4.4. Análisis Técnico de la Descarga y Volcado

#### 1. Sentencia SQL, Campos y Ordenación
- **Sentencia ejecutada:**
  ```sql
  SELECT * FROM Ordenes WHERE Residencia = '<residencia>' ORDER BY Id ASC
  ```
- **Campos:** Recupera todas las columnas mediante `SELECT *`, pero en el bucle de mapeo (`arrOviedo1`, `arrGijon1`, etc.) extrae explícitamente los campos mediante `ObtenerValorCampo(rs, "<NombreCampo>")`.
- **Criterio de ordenación:** `ORDER BY Id ASC`.
  **Hallazgo crítico:** Los registros no se ordenan por `NumOrden` ni por `FechaPeticion`, sino por el autonumérico `Id` de Access (el orden físico o cronológico en el que se insertaron originalmente las filas en Access).

#### 2. Rangos de Limpieza Previa frente a Columnas Escritas

| Residencia | Rangos que Limpia Previamente (`ClearContents`) | Columnas que Escribe con Datos de Access | Columnas Omitidas en la Limpieza Previa |
| :--- | :--- | :--- | :--- |
| **GIJÓN** | `A:T` (Cols 1..20)<br>`V:AA` (Cols 22..27) | `A:T` (Cols 1..20 en `arrGijon1`)<br>`V:AA` (Cols 22..27 en `arrGijon2`) | **Col 21 (`U`, Fórmula)**<br>**Cols 28..29 (`AB:AC`, Fórmulas)**<br>**Col 30 (`AD`, `REPARTIR COSTE  HAB.`)** |
| **SOTO** | `A:P` (Cols 1..16)<br>`S:T` (Cols 19..20)<br>`V:AA` (Cols 22..27) | `A:P` (Cols 1..16 en `arrSoto1`)<br>`S:T` (Cols 19..20 en `arrSoto2`)<br>`V:AA` (Cols 22..27 en `arrSoto3`) | **Cols 17..18 (`Q:R`, Fórmulas)**<br>**Col 21 (`U`, Fórmula)**<br>**Cols 28..29 (`AB:AC`, Fórmulas)**<br>**Col 30 (`AD`, `REPARTIR COSTE  APTO.`)** |
| **OVIEDO** | `A:U` (Cols 1..21)<br>`W:AB` (Cols 23..28) | `A:U` (Cols 1..21 en `arrOviedo1`)<br>`W:AB` (Cols 23..28 en `arrOviedo2`) | **Col 22 (`V`, Fórmula)**<br>**Cols 29..31 (`AC:AE`, Datos Usuario)**<br>**Cols 32..33 (`AF:AG`, Fórmulas)** |

#### 3. Comportamiento de las Columnas a partir de AB (GIJÓN), AD (SOTO) y AG (OVIEDO)
- Las columnas situadas a la derecha de los bloques mapeados **quedan completamente al margen del proceso de sincronización**:
  1. No se limpian en la fase de preparación (`ultimaFila >= filaInicio`).
  2. No se leen de la tabla `Ordenes`.
  3. No se escriben ni actualizan en los bloques matriciales (`arrOviedo`, `arrSoto`, `arrGijon`).
  4. La auto-reparación final de fórmulas únicamente revisa columnas específicas (`AF:AG` en Oviedo, `AB:AC` en Soto y Gijón). Las columnas de datos de usuario no son tocadas.

#### 4. Consecuencias sobre los Datos Locales No Persistidos (`REPARTIR COSTE`, `TAG`, `R`)

Este diseño produce dos efectos colaterales de alto impacto sobre la integridad referencial de los datos de usuario locales:

1. **Desalineación Silenciosa por Desplazamiento de Filas (*Row Drift*):**
   - Cuando se inserta un registro en Access con un `Id` intercalado, o se elimina una orden intermedia en Access, o se filtran/reordenan registros, las filas de datos de Access se vuelcan en Excel secuencialmente desde la fila 2 hacia abajo según `ORDER BY Id ASC`.
   - Sin embargo, las celdas de las columnas no persistidas (`AD` en GIJÓN/SOTO; `AC`, `AD` y `AE` en OVIEDO) **permanecen fijas en sus coordenadas físicas originales**.
   - **Resultado:** El valor de `REPARTIR COSTE`, `TAG` o `R` que pertenecía al expediente de una persona queda asignado silenciosamente a la fila de otra persona totalmente diferente.
2. **Persistencia de Datos Fantasma / Huérfanos (*Ghost Data*):**
   - Si en Access se eliminan órdenes (por ejemplo, 10 expedientes archivados o borrados) y el total de filas descargadas se reduce respecto a la sesión anterior, las filas sobrantes en las columnas no persistidas (`AD` en GIJÓN/SOTO, `AC:AE` en OVIEDO) **no se borran**, permaneciendo visibles como datos huérfanos sin correspondencia con ninguna orden activa.

---

## 5. HECHOS ACREDITADOS Y ASPECTOS NO VERIFICABLES

### 5.1. Hechos Acreditados (Verificados con Código Verbatim)

1. **`EjecutarUpdateCondicionalOrden` valida activamente `RecordsAffected`:**
   El procedimiento ejecuta `cn.Execute sql, affected` y devuelve explícitamente `(affected > 0)`.
2. **Criterio de clave de `EjecutarUpdateCondicionalOrden`:**
   Actualiza exclusivamente por `NumOrden` y `Residencia`. No evalúa el año ni un identificador GUID de instancia.
3. **Mecanismo de Lease Lock:**
   Aplica un cerrojo optimista de 120 segundos basado en `UsuarioModificacion` y `UltimaModificacion`.
4. **Propagación de fracaso en `ActualizarOrden`:**
   Informa al usuario mediante `MsgBox` diferenciando si el expediente no existe o si está bloqueado por otro usuario con el tiempo restante.
5. **Reversión automática en hoja ante conflicto:**
   Si una edición de celda falla en `ActualizarOrdenMultiple`, `SincronizarCambioTargetHaciaAccess` invoca de inmediato a `SincronizarHojaDesdeAccess`, descartando el cambio local y forzando la recarga desde la base de datos.
6. **Nombre real del vaciado integral de alojamientos:**
   El procedimiento real es `BorrarDatosAlojamientoOptimizado` en `BorradoDatosAlojamiento.bas`.
7. **Limpieza integral en `BorrarDatosAlojamientoOptimizado`:**
   Limpia **todos los campos de datos de usuario**, incluyendo explícitamente los cinco campos no persistidos (`AD` en GIJÓN/SOTO; `AC`, `AD`, `AE` en OVIEDO). Protege rigurosamente las columnas de fórmulas calculadas.
8. **Nombre real de la descarga desde Access:**
   El procedimiento real es `SincronizarHojaDesdeAccess` en `modDatabase.bas` (22 llamadas en el proyecto).
9. **Orden de recuperación de `SincronizarHojaDesdeAccess`:**
   Ejecuta `SELECT * FROM Ordenes WHERE Residencia = '...' ORDER BY Id ASC`.
10. **Desconexión física de campos no persistidos en `SincronizarHojaDesdeAccess`:**
    No limpia ni escribe las columnas locales no persistidas (`AD` en GIJÓN/SOTO; `AC:AE` en OVIEDO), provocando que cualquier alteración en el número o secuencia de filas en Access desplace y desalinee los datos locales respecto a los expedientes a los que pertenecían.

### 5.2. Aspectos No Verificables Estáticamente
- No es posible verificar si en algún momento histórico los usuarios finales añadieron columnas personalizadas adicionales a la derecha de la columna `AG` en copias locales no versionadas del libro Excel, puesto que las macros vigentes delimitan expresamente su alcance hasta la columna `AE` (Oviedo) y `AD` (Gijón/Soto).

---

## 6. ESTADO DE GIT ANTES Y DESPUÉS DE LA VERIFICACIÓN

### 6.1. Estado Inicial de Git
- **Comando:** `git status --short`
- **Resultado:** Árbol de trabajo con modificaciones previas existentes no confirmadas asociadas a la sesión activa; ningún archivo adicional modificado.

### 6.2. Estado Final de Git
- **Comando:** `git status --short`
- **Verificación:** Ejecutado tras finalizar las comprobaciones de solo lectura. El estado del árbol de trabajo es exactamente idéntico al inicial. **Cero modificaciones introducidas en el código fuente, libro Excel o base de datos.**

---

## 7. RELACIÓN DE SCRIPTS AUXILIARES UTILIZADOS

En estricto cumplimiento de las restricciones de la Orden 1W, no se creó ningún archivo ni script auxiliar dentro del repositorio `H:\ResidenciaApp\bdas-multiusuario`.

- **Comandos de extracción:** Se ejecutaron mediante invocaciones inline de Python (`python -c "..."`) a través de la terminal estándar.
- **Directorio de persistencia de artefactos e informes:** Todas las evidencias y el presente informe se alojan fuera del repositorio, en el directorio oficial de la sesión:
  `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2/`
