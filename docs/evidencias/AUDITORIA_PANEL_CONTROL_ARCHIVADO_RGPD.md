# Auditoría Técnica, Operativa y Especificación Quirúrgica: Panel de Control y Archivado RGPD

> **Fecha:** 29 de octubre de 2026  
> **Versión Base:** Hito V4 (29/10/2026) — `BDAS_Multiusuario.xlsm` + `Residencia_BE.accdb`  
> **Componentes auditados e intervenidos:** `frmPanelRGPD.frm`, `ModuloArchivadoDatosBDAS.bas`, `ModuloRibbonRGPD.bas`, `BorradoFilasHojasBDAS.bas`

---

## 1. Respuestas Directas a las Cuestiones Planteadas

1. **¿Están operativas y son funcionales las opciones del Panel de Control (Modo Automático/Manual, Días de retención y Guardar Configuración) con la base de datos?**
   * **Sí**, leen y escriben en `CONFIG!B4` (días) y `CONFIG!B5` (modo `"SI"`/`"NO"`) validando permisos de administrador (`modDatabase.ValidarPermisoAdmin`).
   * **Deficiencias corregidas en esta especificación:** No guardaba el libro (`ThisWorkbook.Save`), no auditaba cambios de modo (`B5`) si no cambiaban los días (`B4`), y no registraba el evento en la tabla `LogActividad` de Access (`modDatabase.InsertarLog`).

2. **¿Con la configuración actual se efectúa el borrado automático/manual de los datos de la tabla Excel al cabo de $X$ días?**
   * **Manual (`ArchivarDatosManual`):** **Sí.** Traslada de `Ordenes` a `Ordenes_Historico` en Access y refresca inmediatamente la tabla Excel mediante `modDatabase.SincronizarHojaDesdeAccess`.
   * **Automático (`ArchivarDatosAntiguosAlAbrir`):** Archivaba en Access y actualizaba los `RESUMEN`, pero **no invocaba `modDatabase.SincronizarHojaDesdeAccess`** sobre las hojas `RESIDENCIA GIJÓN/SOTO/OVIEDO`. Se corrige en `[REF-RGPD-03]`.

3. **¿Está previsto en el código actual el borrado automático de los datos de la base de datos al cabo de cierto tiempo ($X$ años)?**
   * **Sí.** `BorradoFilasHojasBDAS.BorrarFilasBDAS()` invoca `modDatabase.PurgarHistoricoRGPD(anosRetencion:=3)`, eliminando físicamente de `Ordenes_Historico` y `BDAS_Historico_Snapshot` los registros con antigüedad $\ge 3\text{ años}$ (Real Decreto 933/2021) y registrando `"PURGA_RGPD"` en `LogActividad`.

---

## 2. Fichas Técnicas de Mejora de Calidad y Robustez

### [REF-RGPD-01] Cabecera `Option Explicit` y Erradicación de Código Muerto en `frmPanelRGPD.frm`
* **Componente / Función afectada:** `frmPanelRGPD.frm` (Cabecera, `EjecutarArchivadoManualConConfirmacion`, bloques comentados de hojas `BDAS` y `ValidarPasswordUsuarioActual`).
* **Diagnóstico del problema:** El formulario carece de `Option Explicit` y conserva procedimientos huérfanos no invocados (`EjecutarArchivadoManualConConfirmacion` y `ValidarPasswordUsuarioActual`, que busca la hoja eliminada `USUARIOS`).
* **Impacto y Justificación:** Garantiza verificación estricta de tipos en compilación y elimina deuda técnica obsoleta sin alterar el archivo binario `.frx`.
* **Nivel de Riesgo de Regresión:** **Bajo**.

### [REF-RGPD-02] Auditoría Integral en Access/LOG y Persistencia de Configuración en `frmPanelRGPD.frm`
* **Componente / Función afectada:** `frmPanelRGPD.frm` $\rightarrow$ `btnGuardar_Click`, `GuardarConfiguracion` y `RegistrarCambioDiasRetencion`.
* **Diagnóstico del problema:**
  1. `btnGuardar_Click` solo evalúa `cambianDias = (dias <> diasAnteriores)`. Si el administrador cambia el modo de `"SI"` a `"NO"` manteniendo los mismos días, el cambio no se audita ni en las hojas `LOG_*` ni en Access.
  2. `RegistrarCambioDiasRetencion` usa `ThisWorkbook.usuarioActual` sin respaldo de `modDatabase.ObtenerUsuarioActual()` y no invoca `modDatabase.InsertarLog`.
  3. `GuardarConfiguracion` modifica `CONFIG!B4:B5` en memoria pero no persiste el libro (`ThisWorkbook.Save`).
* **Impacto y Justificación:** Asegura trazabilidad normativa completa en `LogActividad` (Access) y en las 3 hojas `LOG_*` tanto para cambios de días como de modo, y asegura la persistencia en disco.
* **Nivel de Riesgo de Regresión:** **Bajo**.

### [REF-RGPD-03] Sincronización Visual de Hojas Operativas tras Archivado Automático
* **Componente / Función afectada:** `ModuloArchivadoDatosBDAS.bas` $\rightarrow$ `Public Sub ArchivarDatosAntiguosAlAbrir`.
* **Diagnóstico del problema:** `ArchivarDatosAntiguosAlAbrir` archiva en Access (`ArchivarOrdenesAntiguasBD`) y actualiza las hojas `RESUMEN`, pero no invoca `modDatabase.SincronizarHojaDesdeAccess` para las sedes donde se han archivado órdenes, dejando filas obsoletas en la cuadrícula si la hoja ya estaba cargada.
* **Impacto y Justificación:** Garantiza que las tablas de Excel (`RESIDENCIA GIJÓN`, `SOTO`, `OVIEDO`) queden sincronizadas inmediatamente con Access tras el archivado automático.
* **Nivel de Riesgo de Regresión:** **Bajo**.

### [REF-RGPD-04] Corrección de Celda `CONFIG!B6` en `BuscarBDASActivado` y Limpieza en `ModuloArchivadoDatosBDAS.bas`
* **Componente / Función afectada:** `ModuloArchivadoDatosBDAS.bas` $\rightarrow$ `BuscarBDASActivado`, `RegistrarArchivadoEnLOG` y procedimientos huérfanos (`RegistrarEliminacionFilaEnLOG`, `RegistrarEliminacionMasivaEnLOG`, `ObtenerNombreLOGPorOrigen`).
* **Diagnóstico del problema:**
  1. `BuscarBDASActivado` lee `wsConfig.Range("B5").Value` (que es el modo de archivado automático) en lugar de `B6` (que es el estado de búsqueda en BDAS).
  2. `RegistrarArchivadoEnLOG` recibe los parámetros `gijon`, `soto`, `oviedo` pero no incluye el desglose en la columna D del LOG ni registra el usuario ejecutor real cuando es manual.
  3. `RegistrarEliminacionFilaEnLOG`, `RegistrarEliminacionMasivaEnLOG` y `ObtenerNombreLOGPorOrigen` son código muerto sin llamadores.
* **Impacto y Justificación:** Desacopla el interruptor de búsqueda histórica (`B6`) del modo de archivado automático (`B5`) y mejora el detalle de auditoría en las hojas `LOG_*`.
* **Nivel de Riesgo de Regresión:** **Bajo**.

### [REF-RGPD-05] Eliminación de Doble Confirmación, Validación RBAC y Lectura de `CONFIG!B4` en `ModuloRibbonRGPD.bas`
* **Componente / Función afectada:** `ModuloRibbonRGPD.bas` $\rightarrow$ `RibbonArchivarAhora` y `ObtenerDiasRetencion`.
* **Diagnóstico del problema:**
  1. `RibbonArchivarAhora` carece de validación `modDatabase.ValidarPermisoAdmin`, muestra un mensaje con valor fijo `"90 días"` y referencia obsoleta a `"hoja BDAS"`, y provoca una **doble confirmación** porque acto seguido invoca `ArchivarDatosManual` (que ya cuenta, valida la hoja activa y pide confirmación al usuario).
  2. `ObtenerDiasRetencion` lee `CONFIG!B2` en lugar de `CONFIG!B4`.
* **Impacto y Justificación:** Unifica el flujo del botón del Ribbon con el del Panel de Control, exige permisos de Administrador y corrige la celda de lectura a `B4`.
* **Nivel de Riesgo de Regresión:** **Bajo**.

---

## 3. Bloques Quirúrgicos de Sustitución (`BUSCAR` / `SUSTITUIR POR`)

### OPERACIÓN 01: `frmPanelRGPD.frm` — Adición de `Option Explicit` en cabecera

#### BUSCAR:
```vba
'=================================================================================
' EVENTO: Inicializar formulario
' Se ejecuta automáticamente al abrir el formulario
'=================================================================================
Private Sub UserForm_Initialize()
```

#### SUSTITUIR POR:
```vba
Option Explicit

'=================================================================================
' EVENTO: Inicializar formulario
' Se ejecuta automáticamente al abrir el formulario
'=================================================================================
Private Sub UserForm_Initialize()
```

---

### OPERACIÓN 02: `frmPanelRGPD.frm` — Limpieza de código muerto, auditoría dual (días + modo) en Access/LOG y persistencia en `CONFIG`

#### BUSCAR:
```vba
'=================================================================================
' FUNCIÓN AUXILIAR: Ejecutar Borrado manual con confirmación
'=================================================================================
Private Sub EjecutarArchivadoManualConConfirmacion()
    On Error GoTo ErrorHandler
    
    ' PASO 1: Contar cuántos registros se van a archivar
    Dim totalRegistros As Long
    totalRegistros = ModuloArchivadoDatosBDAS.ContarRegistrosParaArchivar()
    
    If totalRegistros = 0 Then
        Dim msg1 As String
        msg1 = "No hay registros antiguos para archivar." & vbCrLf & vbCrLf
        msg1 = msg1 & "Todos los registros tienen menos de " & txtDiasRetencion.text & " días" & vbCrLf
        msg1 = msg1 & "desde la fecha de salida (finalización del servicio)." & vbCrLf & vbCrLf
        msg1 = msg1 & "Los datos archivados se conservarán en BDAS durante 3 años" & vbCrLf
        msg1 = msg1 & "según Real Decreto 933/2021, de 26 de octubre."
        
        MsgBox msg1, vbInformation, "Borrado - Sin Acción Necesaria"
        Exit Sub
    End If
    
    ' PASO 2: Mostrar resumen y pedir confirmación
    Dim msg2 As String
    Dim msg3 As String
    Dim msg4 As String
    Dim msgCompleto As String
    
    msg2 = "SE VAN A BORRAR " & totalRegistros & " REGISTROS" & vbCrLf & vbCrLf
    msg2 = msg2 & "Registros con más de " & txtDiasRetencion.text & " días desde fecha de salida." & vbCrLf & vbCrLf
'    msg2 = msg2 & "Cada residencia se archivará en su BDAS:" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA GIJÓN -> BDAS GIJÓN" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA SOTO -> BDAS SOTO" & vbCrLf
'    msg2 = msg2 & "  - RESIDENCIA OVIEDO -> BDAS OVIEDO" & vbCrLf & vbCrLf
    
'    msg3 = "DATOS QUE SE ELIMINARÁN de hojas operativas:" & vbCrLf
'    msg3 = msg3 & "  DNI/NIF, Nombre, Teléfono, Dirección, Grabación" & vbCrLf & vbCrLf
'    msg3 = msg3 & "DATOS QUE SE CONSERVAN:" & vbCrLf
'    msg3 = msg3 & "  Nº Orden, Fechas estancia, Importes, Estado pago" & vbCrLf & vbCrLf
'    msg3 = msg3 & "Copia completa guardada en BDAS correspondiente" & vbCrLf
'    msg3 = msg3 & "(conservación: 3 años según RD 933/2021)" & vbCrLf & vbCrLf
    
'    msg4 = "BASE LEGAL:" & vbCrLf
'    msg4 = msg4 & "Real Decreto 933/2021, de 26 de octubre:" & vbCrLf
'    msg4 = msg4 & "  Los datos del registro informático deberán conservarse" & vbCrLf
'    msg4 = msg4 & "  durante un plazo de tres años a contar desde la" & vbCrLf
'    msg4 = msg4 & "  finalización del servicio o prestación contratada." & vbCrLf & vbCrLf
'    msg4 = msg4 & "RGPD (UE 2016/679) - Arts. 5.1. c y 5.1.e" & vbCrLf & vbCrLf
    msg4 = msg4 & "Continuar con el borrado?"
    
    msgCompleto = msg2 & msg3 & msg4
    
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox(msgCompleto, vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Borrado")
    
    If respuesta = vbNo Then
        MsgBox "Operación cancelada por el usuario.", vbInformation, "Borrado"
        Exit Sub
    End If
    
    ' PASO 3: Ejecutar borrado manual
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error en borrado manual: " & Err.Description, vbCritical
End Sub

'=================================================================================
' BOTÓN: Ver BDAS
' Versión mejorada: Detecta qué BDAS existen y adapta el selector
'=================================================================================
'Private Sub btnVerBDAS_Click()
'    On Error GoTo ErrorHandler
    
'    Dim password As String
'    Dim seleccion As String
'    Dim nombreBDAS As String
    
    ' Solicitar contraseña
'    password = InputBox("Acceso a datos archivados (RGPD)" & vbCrLf & vbCrLf & _
'                       "Esta hoja contiene datos personales sensibles." & vbCrLf & _
'                       "Introduzca la contraseña maestra:", _
'                       "Acceso BDAS - Datos Sensibles")
    
    ' Si cancela, salir
'    If password = "" Then Exit Sub
    
    ' Validar contraseña
'    If password <> ModuloConfigSegura.ObtenerClaveMaestra() Then
'        MsgBox "Contraseña incorrecta.", vbCritical, "Acceso Denegado"
'        Exit Sub
'    End If
    
    ' Detectar qué hojas BDAS existen
'    Dim bdasDisponibles As Collection
'    Set bdasDisponibles = DetectarBDASDisponibles()
    
    ' Si no existe ninguna BDAS
'    If bdasDisponibles.Count = 0 Then
'        MsgBox "No existen hojas BDAS." & vbCrLf & vbCrLf & _
'               "No hay datos archivados aún." & vbCrLf & _
'               "Las hojas BDAS se crearán automáticamente cuando se archive por primera vez.", _
'               vbInformation, "Sin BDAS disponibles"
'        Exit Sub
'    End If
    
    ' Si solo existe UNA BDAS, mostrarla directamente
'    If bdasDisponibles.Count = 1 Then
'        nombreBDAS = bdasDisponibles(1)
        
'        MsgBox "Se mostrará la hoja:" & vbCrLf & vbCrLf & _
'               nombreBDAS, vbInformation, "BDAS Disponible"
'    Else
        ' Si existen VARIAS BDAS, mostrar selector dinámico
'        Dim mensajeSelector As String
'        Dim i As Long
        
'        mensajeSelector = "Seleccione qué BDAS desea visualizar:" & vbCrLf & vbCrLf
'        mensajeSelector = mensajeSelector & "Escriba una de las siguientes opciones:" & vbCrLf
        
'        For i = 1 To bdasDisponibles.Count
'            mensajeSelector = mensajeSelector & "  " & i & " - " & bdasDisponibles(i) & vbCrLf
'        Next i
        
'        seleccion = InputBox(mensajeSelector, "Seleccionar BDAS", "1")
        
'        If seleccion = "" Then Exit Sub
        
        ' Validar selección
'        Dim numSeleccion As Long
'        If Not IsNumeric(seleccion) Then
'            MsgBox "Opción no válida.  Debe escribir un número.", vbExclamation, "Selección Inválida"
'            Exit Sub
'        End If
        
'        numSeleccion = CLng(seleccion)
        
'        If numSeleccion < 1 Or numSeleccion > bdasDisponibles.Count Then
'            MsgBox "Opción fuera de rango. Debe escribir un número entre 1 y " & bdasDisponibles.Count, _
'                   vbExclamation, "Selección Inválida"
'            Exit Sub
'        End If
        
        ' Obtener nombre de la BDAS seleccionada
'        nombreBDAS = bdasDisponibles(numSeleccion)
'    End If
    
    ' Verificar que existe la BDAS (doble verificación)
'    Dim wsBDAS As Worksheet
'    On Error Resume Next
'    Set wsBDAS = ThisWorkbook.Worksheets(nombreBDAS)
'    On Error GoTo ErrorHandler
    
'    If wsBDAS Is Nothing Then
'        MsgBox "Error: La hoja " & nombreBDAS & " no pudo ser accedida.", vbCritical, "Error"
'        Exit Sub
'    End If
    
    ' Mostrar mensaje de advertencia
'    MsgBox "DATOS SENSIBLES - RGPD" & vbCrLf & vbCrLf & _
'           "Va a acceder a datos personales archivados de:" & vbCrLf & _
'           nombreBDAS & vbCrLf & vbCrLf & _
'           "Este acceso será registrado en el LOG." & vbCrLf & vbCrLf & _
'           "La hoja se ocultará automáticamente al cambiar de pestaña.", _
'           vbExclamation, "Acceso BDAS"
    
    ' Ocultar formulario y hacer visible BDAS seleccionada
'    Unload Me
'    wsBDAS.visible = xlSheetVisible
'    wsBDAS.Activate
    
    ' Registrar acceso en LOG
'    Call RegistrarAccesoBDASDesdePanel(nombreBDAS)
    
'    Exit Sub
    
'ErrorHandler:
'    MsgBox "Error al acceder a BDAS: " & Err.Description, vbCritical, "Error"
'End Sub

'=================================================================================
' FUNCIÓN AUXILIAR: Detectar qué hojas BDAS existen realmente
' Retorna: Collection con los nombres de las BDAS que existen
'=================================================================================
'Private Function DetectarBDASDisponibles() As Collection
'    Dim disponibles As Collection
'    Set disponibles = New Collection
    
'    Dim ws As Worksheet
'    Dim nombresBDAS As Variant
'    Dim i As Long
    
    ' Lista de nombres posibles de BDAS
'    nombresBDAS = Array("BDAS GIJÓN", "BDAS SOTO", "BDAS OVIEDO")
    
    ' Verificar cuáles existen
'    For i = LBound(nombresBDAS) To UBound(nombresBDAS)
'        On Error Resume Next
'        Set ws = Nothing
'        Set ws = ThisWorkbook.Worksheets(CStr(nombresBDAS(i)))
'        On Error GoTo 0
        
'        If Not ws Is Nothing Then
'            disponibles.Add CStr(nombresBDAS(i))
'        End If
'    Next i
    
'    Set DetectarBDASDisponibles = disponibles
'End Function

'=================================================================================
' FUNCIÓN AUXILIAR: Registrar acceso desde Panel
' ACTUALIZADO: Incluye nombre de BDAS accedida
' CORREGIDO: Ya NO usa Application.UserName (usuario de Windows)
'=================================================================================
'Private Sub RegistrarAccesoBDASDesdePanel(nombreBDAS As String)
'    On Error GoTo ErrorHandler
    
'    Dim wsLOG As Worksheet
'    Dim ultimaFila As Long
'    Dim logPassword As String
'    Dim logSheets As Variant
'    Dim i As Long
'    Dim usuarioActual As String
    
    ' ? CORRECCIÓN: Obtener usuario SOLO de la sesión activa
'    usuarioActual = Trim(ThisWorkbook.usuarioActual)
    
    ' ? VALIDACIÓN: Si está vacío, no debería estar aquí
'    If usuarioActual = "" Then
        ' Esto NO debería ocurrir nunca si el login funciona bien
'        MsgBox "Error crítico: No se puede identificar el usuario de sesión." & vbCrLf & _
'               "Por favor, cierre el libro y vuelva a iniciar sesión.", vbCritical, "Error de Sesión"
'        Debug.Print "ERROR CRÍTICO: usuarioActual vacío en RegistrarAccesoBDASDesdePanel - " & Now
'        Exit Sub
'    End If
    
'    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
'    logSheets = Array("LOG_GIJÓN", "LOG_SOTO", "LOG_OVIEDO")
    
    ' Registrar en todas las hojas LOG
'    For i = LBound(logSheets) To UBound(logSheets)
'        On Error Resume Next
'        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
'        If Not wsLOG Is Nothing Then
'            wsLOG.Unprotect password:=logPassword
'            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
'            wsLOG.Cells(ultimaFila, 1).Value = usuarioActual   ' ? Ahora usa el usuario correcto
'            wsLOG.Cells(ultimaFila, 2).Value = Now
'            wsLOG.Cells(ultimaFila, 3).Value = ""
'            wsLOG.Cells(ultimaFila, 4).Value = "ACCESO " & nombreBDAS
'            wsLOG.Cells(ultimaFila, 5).Value = ""
'            wsLOG.Cells(ultimaFila, 6).Value = ""
'            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
'        End If
'        Set wsLOG = Nothing
'        On Error GoTo ErrorHandler
'    Next i
    
'    Exit Sub
    
'ErrorHandler:
'    Debug.Print "Error en RegistrarAccesoBDASDesdePanel: " & Err.Number & " - " & Err.Description & " (" & Now & ")"
'    Resume Next
'End Sub

'=================================================================================
' BOTÓN: Guardar Configuración
' CORREGIDO: Usa B4 (días), B5 (modo), cierra formulario al terminar
'=================================================================================
Private Sub btnGuardar_Click()
    ' Validar permisos de Administrador (Fase 1 - Paso 3)
    If Not modDatabase.ValidarPermisoAdmin("Modificar Configuración RGPD") Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' VALIDAR días de retención
    If Not IsNumeric(txtDiasRetencion.text) Then
        MsgBox "El valor de 'Días de retención' debe ser un número." & vbCrLf & vbCrLf & _
               "Por favor, introduce un valor numérico.", _
               vbExclamation, "Valor Inválido"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    Dim dias As Long
    dias = CLng(txtDiasRetencion.text)
    
    ' Validar rango
    If dias < 30 Or dias > 365 Then
        Dim msgValidacion As String
        msgValidacion = "Los días de retención deben estar entre 30 y 365." & vbCrLf & vbCrLf
        msgValidacion = msgValidacion & "Valor actual: " & dias & vbCrLf
        msgValidacion = msgValidacion & "Rango permitido: 30 - 365 días" ' & vbCrLf & vbCrLf
'        msgValidacion = msgValidacion & "Nota: Los datos archivados se conservan en BDAS durante" & vbCrLf
'        msgValidacion = msgValidacion & "3 años según Real Decreto 933/2021, de 26 de octubre."
        
        MsgBox msgValidacion, vbExclamation, "Valor Fuera de Rango"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    ' ? Verificar si los días han cambiado
    Dim wsConfig As Worksheet
    Dim diasAnteriores As Long
    Dim cambianDias As Boolean
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    diasAnteriores = CLng(wsConfig.Range("B4").Value)  ' ? B4 no B51
    cambianDias = (dias <> diasAnteriores)
    
    ' ? PASO 1: PEDIR CONFIRMACIÓN (con advertencias si cambian días)
    Dim modoTexto As String
    If optAutomatico.Value Then
        modoTexto = "AUTOMÁTICO al abrir"
    Else
        modoTexto = "SOLO MANUAL"
    End If
    
    Dim respuesta As VbMsgBoxResult
    Dim mensajeConfirmacion As String
    
    mensajeConfirmacion = "¿Guardar cambios en la configuraciónº" & vbCrLf & vbCrLf & _
                          "Modo: " & modoTexto & vbCrLf & _
                          "Días de retención: " & dias & " días"
    
    If cambianDias Then
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "ATENCIÓN - CAMBIO DE CONFIGURACIÓN CRÍTICA:" & vbCrLf & _
                              "  Valor actual: " & diasAnteriores & " días" & vbCrLf & _
                              "  Valor nuevo: " & dias & " días" & vbCrLf & vbCrLf & _
                              "Este cambio afectará al borrado automático RGPD" & vbCrLf & _
                              "y será registrado en el LOG de auditoría."
    Else
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "Los cambios se aplicarán en la próxima sesión."
    End If
    
    respuesta = MsgBox(mensajeConfirmacion, vbQuestion + vbYesNo, "Confirmar Guardado")
    
    If respuesta = vbNo Then
        MsgBox "Cambios descartados.", vbInformation, "Configuración"
        Exit Sub
    End If
    
    '   PASO 2: Validación de Administrador ya realizada (Fase 1 - Paso 3: sin InputBox manual)
    ' ? PASO 3: GUARDAR en CONFIG (solo si pasó todas las validaciones)
    Call GuardarConfiguracion
    
    ' ? PASO 4: Si cambiaron los días, registrar en LOG
    If cambianDias Then
        Call RegistrarCambioDiasRetencion(ThisWorkbook.usuarioActual, diasAnteriores, dias)
    End If
    
    ' ? PASO 5: Mensaje de éxito
    Dim mensajeExito As String
    mensajeExito = "Configuración guardada correctamente." & vbCrLf & vbCrLf & _
                   "Modo: " & modoTexto & vbCrLf & _
                   "Días de retención: " & dias & " días"
    
    If cambianDias Then
        mensajeExito = mensajeExito & vbCrLf & vbCrLf & _
                       "CAMBIO REGISTRADO EN LOG:" & vbCrLf & _
                       "  Usuario: " & ThisWorkbook.usuarioActual & vbCrLf & _
                       "  Cambio: " & diasAnteriores & " a " & dias & " días" & vbCrLf & _
                       "  Fecha: " & Format(Now, "dd/mm/yyyy hh:mm:ss") & vbCrLf & vbCrLf & _
                       "El cambio se aplicará al reiniciar Excel."
    Else
        mensajeExito = mensajeExito & vbCrLf & vbCrLf & _
                       "Los cambios se aplicarán al reiniciar Excel."
    End If
    
    MsgBox mensajeExito, vbInformation, "Configuración Guardada"
    
    ' ? PASO 6: CERRAR EL FORMULARIO
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuración:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub


'=================================================================================
' FUNCIÓN PRIVADA: Guardar configuración en hoja CONFIG
' CORREGIDO: B4 (días), B5 (modo automático)
'=================================================================================
Private Sub GuardarConfiguracion()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim passwordConfig As String
    Dim estabaProtegida As Boolean
    Dim diasGuardar As Long
    
    ' ? 1. Obtener hoja CONFIG
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then
        MsgBox "ERROR: No existe la hoja CONFIG.", vbCritical
        Exit Sub
    End If
    
    ' ? 2. Obtener contraseña y desproteger si es necesario
    passwordConfig = ModuloConfigSegura.ObtenerPasswordConfig()
    estabaProtegida = wsConfig.ProtectContents
    
    If estabaProtegida Then
        wsConfig.Unprotect password:=passwordConfig
    End If
    
    ' ? 3.  GUARDAR MODO AUTOMÁTICO en B5 (no B50)
    If optAutomatico.Value Then
        wsConfig.Range("B5").Value = "SI"
    Else
        wsConfig.Range("B5").Value = "NO"
    End If
    
    ' ? 4.  GUARDAR DÍAS DE RETENCIÓN en B4 (no B51)
    diasGuardar = CLng(txtDiasRetencion.text)
    wsConfig.Range("B4").Value = diasGuardar
    
    ' ? 5.  Agregar etiquetas si no existen
    If wsConfig.Range("A4").Value = "" Then
        wsConfig.Range("A4").Value = "Días de retención RGPD:"
    End If
    If wsConfig.Range("A5").Value = "" Then
        wsConfig.Range("A5").Value = "Borrado automático RGPD:"
    End If
    
    ' ? 6. Reproteger si estaba protegida
    If estabaProtegida Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
    
    ' ? 7.  DEBUG
    Debug.Print "CONFIG guardado - B4=" & wsConfig.Range("B4").Value & " B5=" & wsConfig.Range("B5").Value
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuración:" & vbCrLf & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If estabaProtegida And Not wsConfig Is Nothing Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
End Sub

'=================================================================================
' FUNCIÓN: Validar contraseña del usuario actual
'=================================================================================
Private Function ValidarPasswordUsuarioActual(usuario As String, password As String) As Boolean
    On Error Resume Next
    
    Dim wsUsuarios As Worksheet
    Dim i As Long
    Dim saltUsuario As String
    Dim hashUsuario As String
    
    ValidarPasswordUsuarioActual = False
    
    Set wsUsuarios = ThisWorkbook.Sheets("USUARIOS")
    If wsUsuarios Is Nothing Then Exit Function
    
    For i = 2 To wsUsuarios.Cells(wsUsuarios.Rows.Count, "A").End(xlUp).Row
        If Trim(wsUsuarios.Cells(i, 1).Value) = usuario Then
            saltUsuario = Trim(wsUsuarios.Cells(i, 3).Value)
            hashUsuario = Trim(wsUsuarios.Cells(i, 4).Value)
            
            If ModuloConfigSegura.ValidarPasswordHash(password, saltUsuario, hashUsuario) Then
                ValidarPasswordUsuarioActual = True
                Exit Function
            End If
        End If
    Next i
    
    On Error GoTo 0
End Function

'=================================================================================
' FUNCIÓN: Registrar cambio en LOGs
'=================================================================================
Private Sub RegistrarCambioDiasRetencion(usuario As String, valorAnterior As Long, valorNuevo As Long)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim screenStatus As Boolean
    
    screenStatus = Application.screenUpdating
    Application.screenUpdating = False
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJÓN", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            
            With wsLOG
                .Cells(ultimaFila, 1).Value = usuario
                .Cells(ultimaFila, 2).Value = Now
                .Cells(ultimaFila, 3).Value = "CONFIG"
                .Cells(ultimaFila, 4).Value = "RETENCIÓN DATOS"
                .Cells(ultimaFila, 5).Value = CStr(valorAnterior) & " días"
                .Cells(ultimaFila, 6).Value = CStr(valorNuevo) & " días"
            End With
            
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
    Next i
    
    Application.screenUpdating = screenStatus
    
    On Error GoTo 0
End Sub
```

#### SUSTITUIR POR:
```vba
'=================================================================================
' BOTÓN: Guardar Configuración
' CORREGIDO: Usa B4 (días), B5 (modo), audita ambos en Access y LOGs, persiste libro
'=================================================================================
Private Sub btnGuardar_Click()
    ' Validar permisos de Administrador (Fase 1 - Paso 3)
    If Not modDatabase.ValidarPermisoAdmin("Modificar Configuración RGPD") Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' VALIDAR días de retención
    If Not IsNumeric(txtDiasRetencion.text) Then
        MsgBox "El valor de 'Días de retención' debe ser un número." & vbCrLf & vbCrLf & _
               "Por favor, introduce un valor numérico.", _
               vbExclamation, "Valor Inválido"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    Dim dias As Long
    dias = CLng(txtDiasRetencion.text)
    
    ' Validar rango
    If dias < 30 Or dias > 365 Then
        Dim msgValidacion As String
        msgValidacion = "Los días de retención deben estar entre 30 y 365." & vbCrLf & vbCrLf
        msgValidacion = msgValidacion & "Valor actual: " & dias & vbCrLf
        msgValidacion = msgValidacion & "Rango permitido: 30 - 365 días"
        
        MsgBox msgValidacion, vbExclamation, "Valor Fuera de Rango"
        txtDiasRetencion.SetFocus
        Exit Sub
    End If
    
    ' Verificar valores anteriores en CONFIG (B4 = días, B5 = modo)
    Dim wsConfig As Worksheet
    Dim diasAnteriores As Long
    Dim modoAnteriorRaw As String
    Dim modoAnterior As String
    Dim modoNuevo As String
    Dim cambianDias As Boolean
    Dim cambiaModo As Boolean
    
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    If IsNumeric(wsConfig.Range("B4").Value) And CLng(wsConfig.Range("B4").Value) > 0 Then
        diasAnteriores = CLng(wsConfig.Range("B4").Value)
    Else
        diasAnteriores = 90
    End If
    
    modoAnteriorRaw = UCase$(Trim$(CStr(wsConfig.Range("B5").Value)))
    If modoAnteriorRaw = "SI" Or modoAnteriorRaw = "SÍ" Or modoAnteriorRaw = "TRUE" Or modoAnteriorRaw = "1" Or modoAnteriorRaw = "" Then
        modoAnterior = "SI"
    Else
        modoAnterior = "NO"
    End If
    
    modoNuevo = IIf(optAutomatico.Value, "SI", "NO")
    cambianDias = (dias <> diasAnteriores)
    cambiaModo = (modoNuevo <> modoAnterior)
    
    ' PASO 1: PEDIR CONFIRMACIÓN (con advertencias si cambian días o modo)
    Dim modoTexto As String
    If optAutomatico.Value Then
        modoTexto = "AUTOMÁTICO al abrir"
    Else
        modoTexto = "SOLO MANUAL"
    End If
    
    Dim respuesta As VbMsgBoxResult
    Dim mensajeConfirmacion As String
    
    mensajeConfirmacion = "¿Guardar cambios en la configuración?" & vbCrLf & vbCrLf & _
                          "Modo: " & modoTexto & vbCrLf & _
                          "Días de retención: " & dias & " días"
    
    If cambianDias Or cambiaModo Then
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "ATENCIÓN - CAMBIO DE CONFIGURACIÓN CRÍTICA:"
        If cambianDias Then
            mensajeConfirmacion = mensajeConfirmacion & vbCrLf & _
                                  "  • Retención: " & diasAnteriores & " días -> " & dias & " días"
        End If
        If cambiaModo Then
            mensajeConfirmacion = mensajeConfirmacion & vbCrLf & _
                                  "  • Modo Auto: " & modoAnterior & " -> " & modoNuevo
        End If
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "Este cambio afectará al archivado RGPD y será registrado en el LOG de auditoría."
    Else
        mensajeConfirmacion = mensajeConfirmacion & vbCrLf & vbCrLf & _
                              "No se han detectado cambios respecto a los valores actuales."
    End If
    
    respuesta = MsgBox(mensajeConfirmacion, vbQuestion + vbYesNo, "Confirmar Guardado")
    
    If respuesta = vbNo Then
        MsgBox "Cambios descartados.", vbInformation, "Configuración"
        Exit Sub
    End If
    
    ' PASO 2: GUARDAR en CONFIG y persistir libro
    Call GuardarConfiguracion
    
    ' PASO 3: Si cambiaron los días o el modo, registrar en LOG local y en Access
    Dim usrAudit As String
    usrAudit = modDatabase.ObtenerUsuarioActual()
    If Len(Trim$(usrAudit)) = 0 Then usrAudit = Trim$(ThisWorkbook.usuarioActual)
    If Len(Trim$(usrAudit)) = 0 Then usrAudit = "ADMIN"
    
    If cambianDias Or cambiaModo Then
        Call RegistrarCambioConfigRGPD(usrAudit, diasAnteriores, dias, modoAnterior, modoNuevo)
    End If
    
    ' PASO 4: Mensaje de éxito
    Dim mensajeExito As String
    mensajeExito = "Configuración guardada correctamente." & vbCrLf & vbCrLf & _
                   "Modo: " & modoTexto & vbCrLf & _
                   "Días de retención: " & dias & " días"
    
    If cambianDias Or cambiaModo Then
        mensajeExito = mensajeExito & vbCrLf & vbCrLf & _
                       "CAMBIO REGISTRADO EN LOG Y BASE DE DATOS:" & vbCrLf & _
                       "  Usuario: " & usrAudit & vbCrLf & _
                       "  Retención: " & diasAnteriores & " -> " & dias & " días" & vbCrLf & _
                       "  Modo Auto: " & modoAnterior & " -> " & modoNuevo & vbCrLf & _
                       "  Fecha: " & Format(Now, "dd/mm/yyyy hh:mm:ss")
    End If
    
    MsgBox mensajeExito, vbInformation, "Configuración Guardada"
    
    ' PASO 5: CERRAR EL FORMULARIO
    Unload Me
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuración:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error"
End Sub


'=================================================================================
' FUNCIÓN PRIVADA: Guardar configuración en hoja CONFIG
' CORREGIDO: B4 (días), B5 (modo automático) + guardado persistente del libro
'=================================================================================
Private Sub GuardarConfiguracion()
    On Error GoTo ErrorHandler
    
    Dim wsConfig As Worksheet
    Dim passwordConfig As String
    Dim estabaProtegida As Boolean
    Dim diasGuardar As Long
    
    ' 1. Obtener hoja CONFIG
    Set wsConfig = ThisWorkbook.Sheets("CONFIG")
    
    If wsConfig Is Nothing Then
        MsgBox "ERROR: No existe la hoja CONFIG.", vbCritical
        Exit Sub
    End If
    
    ' 2. Obtener contraseña y desproteger si es necesario
    passwordConfig = ModuloConfigSegura.ObtenerPasswordConfig()
    estabaProtegida = wsConfig.ProtectContents
    
    If estabaProtegida Then
        wsConfig.Unprotect password:=passwordConfig
    End If
    
    ' 3. GUARDAR MODO AUTOMÁTICO en B5
    If optAutomatico.Value Then
        wsConfig.Range("B5").Value = "SI"
    Else
        wsConfig.Range("B5").Value = "NO"
    End If
    
    ' 4. GUARDAR DÍAS DE RETENCIÓN en B4
    diasGuardar = CLng(txtDiasRetencion.text)
    wsConfig.Range("B4").Value = diasGuardar
    
    ' 5. Agregar etiquetas si no existen
    If wsConfig.Range("A4").Value = "" Then
        wsConfig.Range("A4").Value = "Días de retención RGPD:"
    End If
    If wsConfig.Range("A5").Value = "" Then
        wsConfig.Range("A5").Value = "Borrado automático RGPD:"
    End If
    
    ' 6. Reproteger si estaba protegida
    If estabaProtegida Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
    
    ' 7. Persistir cambios en el libro para que no se pierdan al cerrar
    On Error Resume Next
    If Not ThisWorkbook.ReadOnly Then
        ThisWorkbook.Save
    End If
    On Error GoTo ErrorHandler
    
    Debug.Print "CONFIG guardado - B4=" & wsConfig.Range("B4").Value & " B5=" & wsConfig.Range("B5").Value
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al guardar configuración:" & vbCrLf & vbCrLf & Err.Description, vbCritical
    On Error Resume Next
    If estabaProtegida And Not wsConfig Is Nothing Then
        wsConfig.Protect password:=passwordConfig, UserInterfaceOnly:=True
    End If
End Sub

'=================================================================================
' FUNCIÓN: Registrar cambio de configuración RGPD en LOGs locales y en Access
'=================================================================================
Private Sub RegistrarCambioConfigRGPD(ByVal usuario As String, _
                                      ByVal diasAnt As Long, ByVal diasNue As Long, _
                                      ByVal modoAnt As String, ByVal modoNue As String)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim screenStatus As Boolean
    Dim valAntTexto As String
    Dim valNueTexto As String
    
    valAntTexto = CStr(diasAnt) & " días (Auto: " & modoAnt & ")"
    valNueTexto = CStr(diasNue) & " días (Auto: " & modoNue & ")"
    
    screenStatus = Application.screenUpdating
    Application.screenUpdating = False
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJÓN", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            
            With wsLOG
                .Cells(ultimaFila, 1).Value = usuario
                .Cells(ultimaFila, 2).Value = Now
                .Cells(ultimaFila, 3).Value = "CONFIG"
                .Cells(ultimaFila, 4).Value = "CONFIGURACIÓN RGPD"
                .Cells(ultimaFila, 5).Value = valAntTexto
                .Cells(ultimaFila, 6).Value = valNueTexto
            End With
            
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
    Next i
    
    ' Registrar también en la tabla central LogActividad de Access
    Call modDatabase.InsertarLog(usuario, "CONFIG_RGPD", _
                                 "Cambio RGPD: " & valAntTexto & " -> " & valNueTexto)
    
    Application.screenUpdating = screenStatus
    
    On Error GoTo 0
End Sub
```

---

### OPERACIÓN 03: `ModuloArchivadoDatosBDAS.bas` — Sincronización de hojas Excel en `ArchivarDatosAntiguosAlAbrir`, mejora de `RegistrarArchivadoEnLOG` y limpieza de código muerto

#### BUSCAR:
```vba
    If totalArchivados > 0 Then
        ' Actualizar resúmenes si procede
        On Error Resume Next
        If archiGijon > 0 Then Call ModuloResumenGijon.ActualizarResumenGijon
        If archiSoto > 0 Then Call ModuloResumenSoto.ActualizarResumenSoto
        If archiOviedo > 0 Then Call ModuloResumenOviedo.ActualizarResumenOviedo
        On Error GoTo 0
        
        Call RegistrarArchivadoEnLOG(totalArchivados, archiGijon, archiSoto, archiOviedo)
        Call modDatabase.InsertarLog("SISTEMA", "AUTO_ARCHIVADO_RGPD", _
            "Archivados automáticos: " & totalArchivados & " (GIJÓN: " & archiGijon & ", SOTO: " & archiSoto & ", OVIEDO: " & archiOviedo & ")")
            
        MsgBox "Archivado automático completado." & vbCrLf & vbCrLf & _
               "Registros archivados en base histórica: " & totalArchivados & vbCrLf & vbCrLf & _
               "Cumplimiento normativo RGPD actualizado.", vbInformation, "Cumplimiento Normativo"
    End If
```

#### SUSTITUIR POR:
```vba
    If totalArchivados > 0 Then
        ' Sincronizar hojas operativas desde Access y actualizar resúmenes de las sedes afectadas
        On Error Resume Next
        If archiGijon > 0 Then
            Call modDatabase.SincronizarHojaDesdeAccess("GIJÓN", ThisWorkbook.Worksheets("RESIDENCIA GIJÓN"))
            Call ModuloResumenGijon.ActualizarResumenGijon
        End If
        If archiSoto > 0 Then
            Call modDatabase.SincronizarHojaDesdeAccess("SOTO", ThisWorkbook.Worksheets("RESIDENCIA SOTO"))
            Call ModuloResumenSoto.ActualizarResumenSoto
        End If
        If archiOviedo > 0 Then
            Call modDatabase.SincronizarHojaDesdeAccess("OVIEDO", ThisWorkbook.Worksheets("RESIDENCIA OVIEDO"))
            Call ModuloResumenOviedo.ActualizarResumenOviedo
        End If
        On Error GoTo 0
        
        Call RegistrarArchivadoEnLOG(totalArchivados, archiGijon, archiSoto, archiOviedo)
        Call modDatabase.InsertarLog("SISTEMA", "AUTO_ARCHIVADO_RGPD", _
            "Archivados automáticos: " & totalArchivados & " (GIJÓN: " & archiGijon & ", SOTO: " & archiSoto & ", OVIEDO: " & archiOviedo & ")")
            
        MsgBox "Archivado automático completado." & vbCrLf & vbCrLf & _
               "Registros archivados en base histórica: " & totalArchivados & vbCrLf & vbCrLf & _
               "Cumplimiento normativo RGPD actualizado.", vbInformation, "Cumplimiento Normativo"
    End If
```

---

### OPERACIÓN 04: `ModuloArchivadoDatosBDAS.bas` — Desglose de sedes en `RegistrarArchivadoEnLOG` y eliminación de funciones obsoletas sin llamadores

#### BUSCAR:
```vba
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
    logSheets = Array("LOG_GIJÓN", "LOG_SOTO", "LOG_OVIEDO")
    
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
' REGISTRAR ELIMINACIÓN DE FILA EN LOG (similar a BorradoSolicitudes)
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
' OBTENER NOMBRE DE HOJA LOG SEGÚN ORIGEN
'=================================================================================
Private Function ObtenerNombreLOGPorOrigen(ByVal origen As String) As String
    Select Case UCase(origen)
        Case "GIJON"
            ObtenerNombreLOGPorOrigen = "LOG_GIJÓN"
        Case "SOTO"
            ObtenerNombreLOGPorOrigen = "LOG_SOTO"
        Case "OVIEDO"
            ObtenerNombreLOGPorOrigen = "LOG_OVIEDO"
        Case Else
            ObtenerNombreLOGPorOrigen = ""
    End Select
End Function
```

#### SUSTITUIR POR:
```vba
'=================================================================================
' REGISTRAR ARCHIVADO EN LOG (con desglose por residencia)
'=================================================================================
Private Sub RegistrarArchivadoEnLOG(total As Long, gijon As Long, soto As Long, oviedo As Long)
    On Error Resume Next
    
    Dim wsLOG As Worksheet
    Dim ultimaFila As Long
    Dim logPassword As String
    Dim logSheets As Variant
    Dim i As Long
    Dim mensaje As String
    Dim detalleSedes As String
    Dim usrLog As String
    
    usrLog = modDatabase.ObtenerUsuarioActual()
    If Len(Trim$(usrLog)) = 0 Then usrLog = "SISTEMA"
    
    mensaje = "ARCHIVADO RGPD: " & total
    detalleSedes = "GIJÓN: " & gijon & " | SOTO: " & soto & " | OVIEDO: " & oviedo
    
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    logSheets = Array("LOG_GIJÓN", "LOG_SOTO", "LOG_OVIEDO")
    
    For i = LBound(logSheets) To UBound(logSheets)
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(logSheets(i))
        
        If Not wsLOG Is Nothing Then
            wsLOG.Unprotect password:=logPassword
            ultimaFila = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            wsLOG.Cells(ultimaFila, 1).Value = usrLog
            wsLOG.Cells(ultimaFila, 2).Value = Now
            wsLOG.Cells(ultimaFila, 3).Value = mensaje
            wsLOG.Cells(ultimaFila, 4).Value = detalleSedes
            wsLOG.Cells(ultimaFila, 5).Value = ""
            wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True
        End If
        Set wsLOG = Nothing
        On Error GoTo 0
    Next i
End Sub
```

---

### OPERACIÓN 05: `ModuloArchivadoDatosBDAS.bas` — Corrección de celda `CONFIG!B6` en `BuscarBDASActivado`

#### BUSCAR:
```vba
    valorConfig = wsConfig.Range("B5").Value
    
    If IsEmpty(valorConfig) Then
        BuscarBDASActivado = True
    ElseIf UCase(Trim(CStr(valorConfig))) = "SI" Or UCase(Trim(CStr(valorConfig))) = "SÍ" Then
        BuscarBDASActivado = True
    Else
        BuscarBDASActivado = False
    End If
```

#### SUSTITUIR POR:
```vba
    valorConfig = wsConfig.Range("B6").Value
    
    If IsEmpty(valorConfig) Then
        BuscarBDASActivado = False
    ElseIf UCase(Trim(CStr(valorConfig))) = "SI" Or UCase(Trim(CStr(valorConfig))) = "SÍ" Then
        BuscarBDASActivado = True
    Else
        BuscarBDASActivado = False
    End If
```

---

### OPERACIÓN 06: `ModuloRibbonRGPD.bas` — Validación de Administrador, eliminación de doble confirmación en `RibbonArchivarAhora` y lectura de `CONFIG!B4` en `ObtenerDiasRetencion`

#### BUSCAR:
```vba
'=================================================================================
' CALLBACK: Botón "Archivar Ahora" en pestaña RGPD
'=================================================================================
Public Sub RibbonArchivarAhora(control As IRibbonControl)
    On Error GoTo ErrorHandler
    
    ' PASO 1: Contar cuántos registros se van a archivar
    Dim totalRegistros As Long
    totalRegistros = ModuloArchivadoDatosBDAS.ContarRegistrosParaArchivar()
    
    If totalRegistros = 0 Then
        MsgBox "? No hay registros antiguos para archivar." & vbCrLf & vbCrLf & _
               "Todos los registros tienen menos de 90 días desde la fecha de salida.", _
               vbInformation, "Archivado RGPD - Sin Acción Necesaria"
        Exit Sub
    End If
    
    ' PASO 2: Mostrar resumen y pedir confirmación
    Dim respuesta As VbMsgBoxResult
    respuesta = MsgBox("SE VAN A ARCHIVAR " & totalRegistros & " REGISTROS" & vbCrLf & vbCrLf & _
                       "DATOS QUE SE ELIMINARÁN de hojas operativas:" & vbCrLf & _
                       "  • DNI / NIF" & vbCrLf & _
                       "  • Nombre completo" & vbCrLf & _
                       "  • Teléfono" & vbCrLf & _
                       "  • Dirección" & vbCrLf & _
                       "  • Fecha grabación solicitud" & vbCrLf & vbCrLf & _
                       "DATOS QUE SE CONSERVAN (RGPD Art. 5.1.b):" & vbCrLf & _
                       "  • Nº Orden, Fechas estancia, Importes, Estado pago" & vbCrLf & vbCrLf & _
                       "Copia completa guardada en hoja BDAS" & vbCrLf & vbCrLf & _
                       "¿Continuar con el archivado?", _
                       vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar Archivado RGPD")
    
    If respuesta = vbNo Then
        MsgBox "Operación cancelada por el usuario.", vbInformation, "Archivado RGPD"
        Exit Sub
    End If
    
    ' PASO 3: Ejecutar archivado manual
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar archivado:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error en Archivado RGPD"
End Sub
```

#### SUSTITUIR POR:
```vba
'=================================================================================
' CALLBACK: Botón "Archivar Ahora" en pestaña RGPD
'=================================================================================
Public Sub RibbonArchivarAhora(control As IRibbonControl)
    ' Validar permisos de Administrador antes de permitir archivado manual
    If Not modDatabase.ValidarPermisoAdmin("Ejecutar Archivado Manual RGPD") Then Exit Sub

    On Error GoTo ErrorHandler
    
    ' Delegar directamente en ArchivarDatosManual (que valida hoja activa,
    ' lee días dinámicos de CONFIG!B4, cuenta registros y pide confirmación única)
    Call ModuloArchivadoDatosBDAS.ArchivarDatosManual
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error al ejecutar archivado:" & vbCrLf & vbCrLf & _
           Err.Description, vbCritical, "Error en Archivado RGPD"
End Sub
```

---

### OPERACIÓN 07: `ModuloRibbonRGPD.bas` — Corrección de celda `B2` a `B4` en `ObtenerDiasRetencion`

#### BUSCAR:
```vba
    If Not ws Is Nothing Then
        dias = ws.Range("B2").Value
        If IsNumeric(dias) And dias > 0 Then
            ObtenerDiasRetencion = CStr(dias)
            Exit Function
        End If
    End If
```

#### SUSTITUIR POR:
```vba
    If Not ws Is Nothing Then
        dias = ws.Range("B4").Value
        If IsNumeric(dias) And dias > 0 Then
            ObtenerDiasRetencion = CStr(dias)
            Exit Function
        End If
    End If
```

---

## 4. Instrucción Quirúrgica para la IA Implementadora

Copia y entrega el siguiente prompt a tu IA encargada de la implementación:

```text
Actúa como Implementador Técnico de Código de ResidenciaApp siguiendo estrictamente el Protocolo de Implementación Quirúrgica de Cambios.

Archivo objetivo: `H:\Aplicaciones\APP Multiusuario\2_CLIENTES_PUESTOS\BDAS_Multiusuario.xlsm`
Documento de especificación vinculante: `AUDITORIA_PANEL_CONTROL_ARCHIVADO_RGPD.md`

INSTRUCCIONES DE EJECUCIÓN:
1. Genera una copia de seguridad preventiva antes de modificar nada:
   `BDAS_Multiusuario_BACKUP_2026-10-29_PRE_RGPD_V5.xlsm`.
2. Extrae e inspecciona los 3 componentes objetivo:
   - `frmPanelRGPD.frm` (preservando intacto su binario `frmPanelRGPD.frx`)
   - `ModuloArchivadoDatosBDAS.bas`
   - `ModuloRibbonRGPD.bas`
3. Verifica la coincidencia textual exacta (1:1) de las 7 operaciones quirúrgicas (OPERACIÓN 01 a OPERACIÓN 07) especificadas en la Sección 3 de `AUDITORIA_PANEL_CONTROL_ARCHIVADO_RGPD.md`.
4. Si alguna operación no coincide textualmente 1:1, detén la ejecución y repórtalo sin alterar el archivo.
5. Si las 7 operaciones coinciden 1:1, aplícalas conservando la codificación nativa Windows-1252 (ANSI) sin mojibake (GIJÓN, SOTO, OVIEDO, Días, Configuración).
6. Verifica la compilación completa del proyecto VBA (`VBE.CommandBars.FindControl(1, 578).Enabled == False`), confirma que el número total de componentes VBE permanece invariable (172) y genera la copia post-implementación `BDAS_Multiusuario_BACKUP_2026-10-29_POST_RGPD_V5.xlsm`.
7. Emite el Informe Final de Auditoría y Ejecución detallando el resultado de cada operación (OP-01 a OP-07) y el estado de compilación.
```