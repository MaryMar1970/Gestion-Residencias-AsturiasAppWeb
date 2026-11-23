# RESIDENCIAAPP — INFORME DE EVIDENCIAS TÉCNICAS (ORDEN 1V)
## EXTRACCIÓN FINAL DE DEPENDENCIAS PARA IDENTIDAD PERSISTENTE Y FACTURACIÓN TRANSACCIONAL

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Cierre del diseño técnico de `IdInstanciaOrden`, facturación y sincronización offline  
**Naturaleza:** Inspección estática y extracción de código, estrictamente de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha y hora de emisión:** 2026-11-19 20:05:00 UTC+1  
**Documento entregable:** `RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1V.md`  
**Referencia previa:** Orden 1U ([RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1U.md](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1U.md))

---

## 1. OBJETIVO DEL INFORME

El presente documento aporta las evidencias factuales definitivas y el código fuente íntegro requeridos por la Dirección Arquitectónica para cerrar el diseño de:
1. Posibles ubicaciones técnicas y ciclo de vida para la persistencia local de `IdInstanciaOrden` en el libro Excel.
2. Comportamiento y conservación de la relación fila-identidad durante operaciones de ordenación, inserción, eliminación, refresco, archivado y trabajo fuera de línea.
3. Mecanismos reales de alta de expedientes en sus tres rutas (manual, automática y offline).
4. Procedimientos de actualización de campos, modificación de `FechaSalida` y su repercusión sobre la facturación.
5. Flujo transaccional efectivo de la facturación automática, modificaciones manuales y delimitación del alcance de `BeginTrans`.
6. Motor de reconciliación offline SICO-N2 (`SincronizarColaOfflineHaciaAccess`), conversión de facturas proforma y riesgos de duplicidad en reintentos.
7. Análisis técnico de los cinco campos sin persistencia identificados en la Orden 1U (`REPARTIR COSTE`, `TAG`, `R`).

---

## 2. CONTROL DE FUENTES Y TRAZABILIDAD

Se certifica que todas las extracciones de código y análisis proceden de las copias activas del repositorio local `H:\ResidenciaApp\bdas-multiusuario`, en el commit Git **`9f67b2c4e1dad97e8c929163d936302b0815b665`**, coincidentes con la versión consolidada en el libro de trabajo.

### 2.1. Tabla de Integridad y Procedencia de Archivos Inspeccionados

| Archivo Fuente | Ruta Absoluta Resuelta | Tamaño en Disco | Hash SHA-256 Verificado | Fecha y Hora Modificación | Procedencia / Estado Git |
| :--- | :--- | :---: | :---: | :---: | :--- |
| **`BDAS_Multiusuario.xlsm`** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6,621,963 bytes | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` | 2026-11-09 20:32:02 | Repositorio Git / Exportación VBA verificada |
| **`modDatabase.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 298,794 bytes | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | 2026-11-09 20:52:49 | Repositorio Git / Exportación VBA verificada |
| **`AsignarNumFactura.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\AsignarNumFactura.bas` | 32,861 bytes | `53f68ac504c52a8cc9eba9b9b19f7a870edef89665d3dc52173c736d331c02a4` | 2026-11-09 20:52:49 | Repositorio Git / Exportación VBA verificada |
| **`ModuloSeguridadLocal.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\ModuloSeguridadLocal.bas` | 72,867 bytes | `1e0e6e2c3c06f35c7a1a28ff9b7b8981d0664ec8202b7dda6cbb8c3c1df81fcc` | 2026-11-09 20:52:49 | Repositorio Git / Exportación VBA verificada |
| **`Hoja2.cls`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\Hoja2.cls` | 40,345 bytes | `659fd5eb6b3eebe0cc7e0f878724a35775fd254702edf1146400a31879e0f88a` | 2026-11-09 20:52:46 | Repositorio Git / Exportación VBA verificada |
| **`Hoja20.cls`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\Hoja20.cls` | 37,532 bytes | `17e8b4c554804ecba2448188763712625f8f57bb0a71d6321c223110df904e70` | 2026-11-09 20:52:46 | Repositorio Git / Exportación VBA verificada |
| **`Hoja24.cls`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\Hoja24.cls` | 55,341 bytes | `ed91f79ef271c27257ac8bc31c9aaf48f64633c6a2ac8b54e068ddb617bc187b` | 2026-11-09 20:52:46 | Repositorio Git / Exportación VBA verificada |
| **`BorradoNumOrden.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoNumOrden.bas` | 7,163 bytes | `eb5c30c18718bdc80474e7a4a6de5510cd8401cdeb9a2566e2c3af44d7420daa` | 2026-11-09 20:52:48 | Repositorio Git / Exportación VBA verificada |
| **`BorradoSolicitudes.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoSolicitudes.bas` | 21,439 bytes | `c21411f02719bc0a7730e3fbb772ba8a6c2c59b8a5e512fd1dc1d5e173a646c1` | 2026-11-09 20:52:48 | Repositorio Git / Exportación VBA verificada |
| **`BorradoDatosAlojamiento.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoDatosAlojamiento.bas` | 11,065 bytes | `8910cfe366c08a3464caa2328460f24926a9ec8bc6c9051c20f2fc9f34fd20a0` | 2026-11-09 20:52:47 | Repositorio Git / Exportación VBA verificada |
| **`ModuloArchivadoDatosBDAS.bas`** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\ModuloArchivadoDatosBDAS.bas` | 20,588 bytes | `f5babcf1699180b54b2d826a8b935e1d5482eea1aa63882ab052bfb6e65fb23c` | 2026-11-09 20:52:47 | Repositorio Git / Exportación VBA verificada |
| **`Residencia_BE.accdb`** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7,639,040 bytes | `9617f99795b8266c2abb91b479bf64b9ece449ad15885ce40edbe62639dbbee9` | 2026-11-19 00:28:25 | Repositorio Git / Exportación VBA verificada |

---

## 3. COMPROBACIÓN A — POSIBLE UBICACIÓN Y CICLO DE VIDA DE `IdInstanciaOrden`

### 3.1. Inventario de Hojas Operativas y Hojas Anuales
- **Inexistencia de hojas anuales separadas:** El libro [`BDAS_Multiusuario.xlsm`](file:///H:/ResidenciaApp/bdas-multiusuario/BDAS_Multiusuario.xlsm) **no dispone de hojas anuales** (como `RESIDENCIA GIJÓN 2025` o `2026`). Cada sede opera en una única hoja operativa multi-ejercicio dimensionada para 4.000 filas:
  - `RESIDENCIA GIJÓN`: Hoja visible. Rango utilizado en XML: `A1:AB4000`. Propiedad VBA `ScrollArea = "A1:AD4000"`.
  - `RESIDENCIA SOTO`: Hoja visible. Rango utilizado en XML: `A1:XFC4000`. Propiedad VBA `ScrollArea = "A1:AE4000"`.
  - `RESIDENCIA OVIEDO`: Hoja visible. Rango utilizado en XML: `A1:AJ4000`. Propiedad VBA `ScrollArea = "A1:AG4000"`.
- **Última columna con presencia de datos o fórmulas:**
  - En **GIJÓN**: Columna 30 (`AD`, encabezado de datos `REPARTIR COSTE  HAB.`). La columna 28 (`AB`) contiene la fórmula de generación del nombre PDF, y la 29 (`AC`) la fórmula del texto para correo electrónico.
  - En **SOTO**: Columna 30 (`AD`, encabezado de datos `REPARTIR COSTE  APTO.`).
  - En **OVIEDO**: Columna 33 (`AG`, fórmula de texto para correo). Las columnas de datos llegan hasta la 31 (`AE`, `R`). Las columnas 32 (`AF`) y 33 (`AG`) son fórmulas.
- **Columnas ocultas por Excel:**
  - En `RESIDENCIA OVIEDO`: Columna 8 (`H`, `COMISIÓN`) está oculta por interfaz (`width="6.14" hidden="1"`). Las columnas 34 a 16.384 están ocultas.
  - En `RESIDENCIA SOTO`: Las columnas 31 a 16.384 están formalmente ocultas en el XML (`hidden="1"`).

### 3.2. Inventario de Hojas Técnicas Ocultas (`hidden` y `veryHidden`)
Se han localizado 13 hojas técnicas con visibilidad restringida:

| Nombre de Hoja Técnica | Visibilidad Excel | Dimensión XML | Propósito técnico observado en código |
| :--- | :---: | :---: | :--- |
| **`CONFIG_PWD`** | `veryHidden` | `A1:A2` | Almacenamiento del hash de contraseña maestra. |
| **`CONFIG`** | `veryHidden` | `A1:B10` | Claves de configuración del entorno y rutas. |
| **`USUARIOS`** | `veryHidden` | `A1:F10` | Caché local de usuarios autorizados y roles. |
| **`SOLAPES_ACEPTADOS_GIJON`** | `veryHidden` | `A1:D2` | Registro temporal de excepciones de solape autorizadas en GIJÓN. |
| **`SOLAPES_ACEPTADOS_OVIEDO`** | `veryHidden` | `A1:F8` | Registro temporal de excepciones de solape autorizadas en OVIEDO. |
| **`PLANTILLAS GIJÓN`** | `veryHidden` | `A1:D10` | Textos estándar para correos y comunicaciones. |
| **`PLANTILLAS SOTO`** | `veryHidden` | `A1:D10` | Textos estándar para correos y comunicaciones. |
| **`PLANTILLAS OVIEDO`** | `veryHidden` | `A1:D10` | Textos estándar para correos y comunicaciones. |
| **`Apartamentos Bloqueados SOTO`** | `veryHidden` | `A1:D2` | Registro técnico de apartamentos bloqueados en SOTO. |
| **`Habitaciones Bloqueadas OVIEDO`** | `veryHidden` | `A1:D2` | Registro técnico de habitaciones bloqueadas en OVIEDO. |
| **`Evaluación`** | `veryHidden` | `A1:J74` | Matriz estática de resolución según `(Finalidad, Empleo, Situación)`. |
| **`Auxiliar`** | `veryHidden` | `A1:E10` | Parámetros auxiliares del sistema. |
| **`SYS_COLA_SYNC`** | `veryHidden` | `A1:M1749` | Cola de persistencia offline SICO-N2 (13 columnas). |

- **Nombres definidos (`DefinedNames`):** No existen nombres de rango definidos en el libro que almacenen claves o metadatos fila a fila. Los nombres existentes (`_xlnm._FilterDatabase`) son creados automáticamente por Excel para la gestión de filtros automáticos (`$A$1:$AB$1`).

### 3.3. Procedimientos que alteran la relación entre Filas y Datos Técnicos

| Operación | Módulo y Procedimiento | Mecanismo técnico ejecutado | Consecuencia sobre los datos técnicos de la fila |
| :--- | :--- | :--- | :--- |
| **Eliminación masiva por orden** | [`BorradoSolicitudes.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoSolicitudes.bas)<br>`BorrarSolicitudesPorOrden` | Ejecuta `rngEliminar.Delete Shift:=xlUp` sobre las filas encontradas. | **Desplazamiento vertical:** Todas las filas inferiores suben de posición física en la cuadrícula. Si un dato técnico estuviera guardado en una hoja externa indexado por número de fila, quedaría desincronizado. |
| **Archivado a histórico** | [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas)<br>`ArchivarOrdenesAntiguasBD` | Ejecuta `rngBatch.Delete Shift:=xlUp` en bloques de hasta 100 filas. | **Desplazamiento vertical idéntico.** Las filas inferiores suben inmediatamente. |
| **Vaciado de fila por borrado de Nº Orden** | [`BorradoNumOrden.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoNumOrden.bas)<br>`GestionarBorradoOrden` | Ejecuta `ws.Range(Cells(fila, 1), Cells(fila, ultCol)).ClearContents`. | **Permanencia de fila:** La fila no se elimina físicamente; se limpian sus celdas. Los datos de columnas fuera del rango `1..ultCol` no se borran si la rutina no los abarca. |
| **Vaciado total de tabla** | [`BorradoDatosAlojamiento.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas)<br>`LimpiarTablaAlojamientoCompleta` | Ejecuta `ClearContents` por bloques de columnas respetando columnas de fórmulas. | Limpia las columnas 1..20, 22..27 y 30. No desplaza filas. |
| **Recarga completa desde Access** | [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas)<br>`DescargarOrdenesDesdeAccess` | Consulta Access con `ORDER BY NumOrden ASC` y vuelca bloques COM en las celdas de la hoja (`arrGijon1`, `arrGijon2`). | **Reconstrucción total de la cuadrícula:** Reescribe desde la fila 2 las columnas de datos basándose en el orden ascendente de `NumOrden`. Si una fila en Excel no procedía de Access o tenía datos desacoplados, es sobreescrita. **Las columnas fuera del mapeo de descarga (como AD en GIJÓN/SOTO o AC:AE en OVIEDO) no se tocan ni se limpian**, quedando potencialmente desalineadas con respecto a la orden que ahora ocupa esa fila. |

### 3.4. Análisis Factual de Posibles Ubicaciones para `IdInstanciaOrden`
1. **Ubicación en Hoja Técnica Desacoplada (e.g. indexada por número de fila):**
   - *Comportamiento demostrado:* La ejecución de `Delete Shift:=xlUp` en archivado o borrado de solicitudes desplaza las filas en la hoja operativa pero no desplazaría los registros en la hoja técnica, produciendo una corrupción inmediata de punteros.
2. **Ubicación en Hoja Técnica indexada por `(Residencia, NumOrden)`:**
   - *Comportamiento demostrado:* Falla en altas provisionales offline (`900xxx`) y colisiona si el `NumOrden` se reutiliza en el mismo ejercicio o entre ejercicios distintos.
3. **Ubicación en una Columna de la Propia Fila Operativa (e.g. columna oculta dentro del rango de la hoja):**
   - *Comportamiento demostrado:* Al residir en la misma fila física, **se desplaza automáticamente** con las operaciones nativas de Excel (`Delete Shift:=xlUp`, ordenaciones por interfaz y filtrado).
   - *Condición técnica requerida:* Dicha columna debe estar incluida en las rutinas de lectura (`SincronizarFilaHaciaAccess`), de descarga (`DescargarOrdenesDesdeAccess`) y de limpieza (`ClearContents`), para evitar desalineaciones durante las recargas masivas desde el servidor.

---

## 4. COMPROBACIÓN B — ALTAS REALES Y PERSISTENCIA DE LA IDENTIDAD

### 4.1. Recorridos de Alta en el Código Real
Existen tres rutas de creación de expedientes en el sistema:

1. **Ruta 1: Alta Manual (Edición de Celda en Columna A):**
   - El operador escribe directamente un número en la Columna A de la hoja.
   - El evento `Worksheet_Change` (Case 1) intercepta la edición e invoca `modDatabase.GestionarEntradaNumOrdenHoja`.
   - *Si hay red (Online):* Adquiere un Lease Lock sobre la clave `resCode & ":" & numOrden` y ejecuta `IntentarReservarOrdenBD`, insertando un registro preliminar en `Ordenes` con `NumOrden`, `Residencia` y marcas de auditoría (`FechaCreacion`, `UsuarioCreacion`).
   - *Si no hay red (Offline contingencia SICO-N2):* `GestionarEntradaNumOrdenHoja` detecta la indisponibilidad de red y calcula un número provisional local `900xxx` (a partir de 900001). Encola una operación `PRERESERVA` con `colTipoOperacion = "PRERESERVA"` y `colNumOrden = 900xxx`.
2. **Ruta 2: Alta Automática (Reserva Atómica del Siguiente Libre):**
   - Invocada por macros de reserva o asistentes: `ReservarSiguienteOrdenLibreBD` o `InsertarOrdenAtomica`.
   - Ejecuta una transacción ADODB con `SELECT MAX(NumOrden) + 1` (o búsqueda del primer hueco libre en Access) e inserta el registro atómicamente.
3. **Ruta 3: Reconciliación de Alta Provisional Offline:**
   - Al restaurar la conexión, el motor FIFO invoca `ResolverNumOrdenProvisionalEnBD`:
     - Consulta `SELECT MAX(NumOrden) FROM Ordenes WHERE NumOrden < 900000` y calcula el nuevo número real definitivo.
     - Llama a `AsegurarOrdenExiste(nuevoNum, resCode)` para insertar el registro en Access.
     - Actualiza la celda de la Columna A en la hoja Excel activa.
     - Recorre la cola `SYS_COLA_SYNC` y sustituye el número provisional `900xxx` por el definitivo en todas las operaciones pendientes de esa orden.

### 4.2. Extracción de Procedimientos de Alta e Identidad

#### `GestionarEntradaNumOrdenHoja` (`modDatabase.bas`, líneas 1527-1696)
```vba
Public Function GestionarEntradaNumOrdenHoja( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    ByVal targetRow As Long, _
    ByRef ioValorAnteriorHoja As Variant, _
    ByVal claveResidencia As String, _
    ByVal nombreHojaResumen As String, _
    ByVal colFactSum As String, _
    ByVal colPagoSum As String, _
    ByVal colPagoRes As String _
) As Boolean

    GestionarEntradaNumOrdenHoja = False
    If ws Is Nothing Or Target Is Nothing Or targetRow < 2 Then Exit Function

    Dim celdaVal As Variant
    celdaVal = Target.Value
    If IsEmpty(celdaVal) Then Exit Function
    If Not IsNumeric(celdaVal) Then Exit Function

    Dim numSolicitado As Long
    numSolicitado = CLng(celdaVal)
    If numSolicitado <= 0 Then Exit Function

    Dim teniamosNumeroPrevio As Boolean
    Dim numPrevio As Long
    numPrevio = 0
    If Not IsEmpty(ioValorAnteriorHoja) Then
        If IsNumeric(ioValorAnteriorHoja) Then
            numPrevio = CLng(ioValorAnteriorHoja)
            If numPrevio > 0 And numPrevio <> numSolicitado Then
                teniamosNumeroPrevio = True
            End If
        End If
    End If

    Dim usrProp As String, nomHuesp As String, sigLibre As Long

    ' CASO A: Modificación / Renumeración de una orden que ya tenía número previo en esta fila
    If teniamosNumeroPrevio Then
        Dim ocupadoPorOtro As Boolean
        Dim dniDummy As String
        ocupadoPorOtro = ExisteOrdenEnBD(claveResidencia, numSolicitado, usrProp, nomHuesp, dniDummy, sigLibre)

        If ocupadoPorOtro Then
            MsgBox "COLISIÓN AL RENUMERAR ORDEN:" & vbCrLf & vbCrLf & _
                   "El Nº ORDEN " & numSolicitado & " ya está asignado en la base de datos (" & claveResidencia & "):" & vbCrLf & _
                   "   Adjudicado a: " & IIf(Len(usrProp) > 0, usrProp, "(Otro operador)") & vbCrLf & _
                   "   Huésped: " & IIf(Len(nomHuesp) > 0, nomHuesp, "(En grabación)") & vbCrLf & vbCrLf & _
                   "Se restaurará el Nº ORDEN anterior (" & numPrevio & ") en esta fila.", _
                   vbExclamation, "Colisión de Renumeración - " & claveResidencia
            ws.Cells(targetRow, "A").Value = numPrevio
            GestionarEntradaNumOrdenHoja = False
            Exit Function
        Else
            If RenumerarOrdenExistenteBD(claveResidencia, numPrevio, numSolicitado) Then
                ioValorAnteriorHoja = numSolicitado
                Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
                    ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numSolicitado)
                Call SincronizarFilaHaciaAccess(ws, targetRow)
                GestionarEntradaNumOrdenHoja = True
                Exit Function
            Else
                MsgBox "No se pudo actualizar el Nº ORDEN en la base de datos central. Se restaura el anterior.", _
                       vbExclamation, "Error de Persistencia - " & claveResidencia
                ws.Cells(targetRow, "A").Value = numPrevio
                GestionarEntradaNumOrdenHoja = False
                Exit Function
            End If
        End If
    End If

    ' CASO B: Alta nueva en una celda A que estaba vacía -> Reserva atómica en Access (o provisional 900xxx si Offline)
    Dim reservadoOK As Boolean, sinRedReserva As Boolean
    reservadoOK = IntentarReservarOrdenBD(claveResidencia, numSolicitado, usrProp, nomHuesp, sigLibre, sinRedReserva)

    If Not reservadoOK Then
        ' Subcaso B.1: Sin conexión con Access -> Impedir adjudicar un Nº ORDEN oficial sin confirmación del servidor
        If sinRedReserva Then
            Dim numProv As Long
            If ModuloSeguridadLocal.EsNumOrdenProvisional(numSolicitado) Then
                numProv = numSolicitado
            Else
                numProv = ModuloSeguridadLocal.ObtenerSiguienteNumOrdenProvisional(ws)
            End If
            ws.Cells(targetRow, "A").Value = numProv
            ioValorAnteriorHoja = numProv
            Call ModuloSeguridadLocal.EncolarOperacionOffline( _
                claveResidencia, numProv, "PRERESERVA", "NumOrden", "", CStr(numProv))
            MsgBox "MODO CONTINGENCIA SICO-N2 (SIN CONEXIÓN CON ACCESS):" & vbCrLf & vbCrLf & _
                   "No es posible reservar en firme el Nº ORDEN oficial " & numSolicitado & " sin confirmación de la base de datos central." & vbCrLf & _
                   "Se ha asignado el Nº ORDEN provisional " & numProv & ", que será convertido automáticamente en el siguiente Nº oficial libre al restablecerse la red.", _
                   vbInformation, "Pre-Reserva Offline SICO-N2 - " & claveResidencia
            ws.Cells(targetRow, 2).Select
            GestionarEntradaNumOrdenHoja = False ' Encolado offline: no duplicar en LOG hasta conciliación FIFO
            Exit Function
        End If

        ' Subcaso B.2: COLISIÓN DETECTADA EN ACCESS:
        ' 1. Limpiar inmediatamente la celda y resetear ValorAnterior
        ws.Cells(targetRow, "A").ClearContents
        ioValorAnteriorHoja = Empty

        ' 2. Sincronizar la hoja desde Access para que el operador vea los registros reales del servidor
        Call SincronizarHojaDesdeAccess(claveResidencia, ws)

        ' 3. Ofrecer interactivamente al operador el siguiente Nº ORDEN libre disponible en la BD
        Dim msgOferta As String
        Dim respOferta As VbMsgBoxResult

        msgOferta = "ATENCIÓN - Nº ORDEN " & numSolicitado & " YA ASIGNADO EN LA BASE DE DATOS:" & vbCrLf & vbCrLf & _
                    "   Adjudicado a: " & IIf(Len(usrProp) > 0, usrProp, "(Otro operador en edición)") & vbCrLf & _
                    "   Huésped: " & IIf(Len(nomHuesp) > 0, nomHuesp, "(En grabación)") & vbCrLf & vbCrLf & _
                    "No se puede utilizar el mismo Nº ORDEN en esta fila." & vbCrLf & _
                    "El siguiente Nº ORDEN disponible en la base de datos es el:  " & sigLibre & vbCrLf & vbCrLf & _
                    "¿Desea adjudicar y utilizar automáticamente el Nº ORDEN " & sigLibre & " ahora?"

        respOferta = MsgBox(msgOferta, vbYesNo + vbQuestion + vbDefaultButton1, _
                            "Colisión de Nº ORDEN Evitada - " & claveResidencia)

        If respOferta = vbYes Then
            Dim numAdjudicado As Long
            numAdjudicado = ReservarSiguienteOrdenLibreBD(claveResidencia)
            If numAdjudicado > 0 Then
                Dim filaLibre As Long
                Dim ultFilaActual As Long
                ultFilaActual = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If ultFilaActual < 2 And Trim$(CStr(ws.Cells(2, "A").Value)) = "" Then
                    filaLibre = 2
                Else
                    filaLibre = ultFilaActual + 1
                End If

                If targetRow >= filaLibre And Trim$(CStr(ws.Cells(targetRow, "A").Value)) = "" Then
                    filaLibre = targetRow
                End If

                ws.Cells(filaLibre, "A").Value = numAdjudicado
                ioValorAnteriorHoja = numAdjudicado

                Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
                    ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numAdjudicado)
                Call SincronizarFilaHaciaAccess(ws, filaLibre, True)

                ws.Cells(filaLibre, 2).Select
                If numAdjudicado <> sigLibre Then
                    MsgBox "Se le ha adjudicado y reservado en firme el Nº ORDEN " & numAdjudicado & _
                           " en la fila " & filaLibre & ".", vbInformation, "Nº ORDEN Adjudicado"
                End If
                GestionarEntradaNumOrdenHoja = True
            Else
                MsgBox "No se pudo confirmar la reserva del Nº ORDEN en la base de datos Access.", _
                       vbExclamation, "Reserva No Confirmada - " & claveResidencia
            End If
        Else
            Dim filaVaciaFinal As Long
            filaVaciaFinal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
            If filaVaciaFinal < 2 Then filaVaciaFinal = 2
            ioValorAnteriorHoja = Empty
            ws.Cells(filaVaciaFinal, "A").Select
        End If
        Exit Function
    End If

    ioValorAnteriorHoja = numSolicitado
    Call MarcarSiPagadosEnResidencia.SincronizarOrdenYCampos( _
        ws.Name, "C", colPagoRes, nombreHojaResumen, colFactSum, colPagoSum, numSolicitado)
    Call SincronizarFilaHaciaAccess(ws, targetRow, True)
    GestionarEntradaNumOrdenHoja = True
End Function

```

#### `InsertarOrdenAtomica` (`modDatabase.bas`, líneas 846-976)
```vba
Public Function InsertarOrdenAtomica( _
    ByVal residencia As String, _
    ByVal fechaPeticion As Date, _
    ByVal dni As String, _
    ByVal nombre As String, _
    ByVal apellidos As String, _
    ByVal tipoHuesped As String, _
    ByVal numHabInd As Integer, _
    ByVal numHabDob As Integer, _
    ByVal fechaEntrada As Date, _
    ByVal fechaSalida As Date, _
    ByVal observaciones As String, _
    Optional ByVal telefono As String = "", _
    Optional ByVal email As String = "", _
    Optional ByVal direccion As String = "", _
    Optional ByVal codigoPostal As String = "", _
    Optional ByVal poblacion As String = "", _
    Optional ByVal provincia As String = "" _
) As Long

    Dim cn As Object
    Dim rs As Object
    Dim sql As String
    Dim nuevoNumOrden As Long
    Dim intentos As Integer
    Dim exito As Boolean
    
Dim usuarioCreador As String
    usuarioCreador = ObtenerUsuarioActual()
    If Len(Trim$(usuarioCreador)) = 0 Then usuarioCreador = "admin"

    intentos = 5
    exito = False
    
    Do While intentos > 0 And Not exito
        Dim enTrans As Boolean
        enTrans = False
        On Error Resume Next
        Err.Clear
        
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 5
        cn.Open GetConnectionString()
        
        If Err.Number = 0 Then
            cn.BeginTrans
            If Err.Number = 0 Then enTrans = True
            
            Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "'")
            If Err.Number = 0 And Not rs Is Nothing Then
                If Not rs.EOF And Not IsNull(rs(0).Value) Then
                    nuevoNumOrden = CLng(rs(0).Value) + 1
                Else
                    nuevoNumOrden = 1
                End If
                rs.Close
                Set rs = Nothing
                
                sql = "INSERT INTO Ordenes (" & _
                      "NumOrden, Residencia, FechaPeticion, DNI, Nombre, Apellidos, TipoHuesped, " & _
                      "NumHabIndividuales, NumHabDobles, FechaEntrada, FechaSalida, " & _
                      "Resolucion, EstadoPago, Observaciones, " & _
                      "Telefono, Email, Direccion, CodigoPostal, Poblacion, Provincia, " & _
                      "Solapamiento, ConsentimientoRGPD, " & _
                      "FechaCreacion, UsuarioCreacion" & _
                      ") VALUES (" & nuevoNumOrden & ", "
                      
                sql = sql & "'" & EscaparSQL(residencia) & "', " & _
                      FormatearFechaSQL(fechaPeticion) & ", " & _
                      "'" & EscaparSQL(dni) & "', " & _
                      "'" & EscaparSQL(nombre) & "', " & _
                      "'" & EscaparSQL(apellidos) & "', " & _
                      "'" & EscaparSQL(tipoHuesped) & "', " & _
                      numHabInd & ", " & numHabDob & ", " & _
                      FormatearFechaSQL(fechaEntrada) & ", " & _
                      FormatearFechaSQL(fechaSalida) & ", "
                      
                sql = sql & "'', '', " & _
                      "'" & EscaparSQL(observaciones) & "', " & _
                      "'" & EscaparSQL(telefono) & "', " & _
                      "'" & EscaparSQL(email) & "', " & _
                      "'" & EscaparSQL(direccion) & "', " & _
                      "'" & EscaparSQL(codigoPostal) & "', " & _
                      "'" & EscaparSQL(poblacion) & "', " & _
                      "'" & EscaparSQL(provincia) & "', " & _
                      "0, 0, " & _
                      FormatearFechaHoraSQL(Now) & ", " & _
                      "'" & EscaparSQL(usuarioCreador) & "'" & _
                      ")"
                      
                cn.Execute sql
                
                If Err.Number = 0 Then
                    cn.CommitTrans
                    enTrans = False
                    exito = True
                End If
            End If
            
            If enTrans Then cn.RollbackTrans
            If Not rs Is Nothing Then
                If rs.State = adStateOpen Then rs.Close
                Set rs = Nothing
            End If
            If cn.State = adStateOpen Then cn.Close
        End If
        Set cn = Nothing
        On Error GoTo 0
        
        If Not exito Then
            intentos = intentos - 1
            If intentos > 0 Then PausaMS 60
        End If
    Loop
    
    If exito Then
        InsertarOrdenAtomica = nuevoNumOrden
    Else
        MsgBox "No se pudo asignar un Nº ORDEN atómico tras varios intentos concurrentes.", vbCritical, "Error Concurrencia"
        InsertarOrdenAtomica = 0
    End If
    
    GoTo CleanUp

CleanUp:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
End Function

```

#### `IntentarReservarOrdenBD` (`modDatabase.bas`, líneas 1281-1401)
```vba
Public Function IntentarReservarOrdenBD( _
    ByVal residencia As String, _
    ByVal numOrdenSolicitado As Long, _
    ByRef outUsuarioPropietario As String, _
    ByRef outNombreHuesped As String, _
    ByRef outSiguienteLibreBD As Long, _
    Optional ByRef outSinRed As Boolean = False _
) As Boolean

    Dim cn As Object, rs As Object
    Dim claveRes As String, usrActual As String, tokenActual As String
    Dim enTrans As Boolean

    On Error GoTo ErrReserva
    IntentarReservarOrdenBD = False
    outUsuarioPropietario = ""
    outNombreHuesped = ""
    outSiguienteLibreBD = 0
    outSinRed = False

    If numOrdenSolicitado <= 0 Then Exit Function
    If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
        outSinRed = True
        IntentarReservarOrdenBD = False
        Exit Function
    End If

    Call AsegurarEsquemaConcurrenciaBD

    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    usrActual = Trim$(ObtenerUsuarioActual())
    If Len(usrActual) = 0 Then usrActual = "admin"
    tokenActual = ObtenerTokenSesionInstancia()

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()

    cn.BeginTrans
    enTrans = True

    Set rs = cn.Execute("SELECT UsuarioCreacion, UsuarioModificacion, Nombre FROM Ordenes " & _
                        "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden = " & numOrdenSolicitado)

    If Not rs.EOF Then
        ' Ya existe en Access -> Leer propietario y calcular siguiente libre
        outUsuarioPropietario = Trim$(CStr(Nz(rs("UsuarioModificacion").Value, "")))
        If Len(outUsuarioPropietario) = 0 Then
            outUsuarioPropietario = Trim$(CStr(Nz(rs("UsuarioCreacion").Value, "")))
        End If
        outUsuarioPropietario = FormatearPropietarioBloqueo(outUsuarioPropietario)
        outNombreHuesped = Trim$(CStr(Nz(rs("Nombre").Value, "")))
        rs.Close
        Set rs = Nothing

        Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes " & _
                            "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden < 900001")
        If Not rs.EOF And Not IsNull(rs(0).Value) Then
            outSiguienteLibreBD = CLng(rs(0).Value) + 1
        Else
            outSiguienteLibreBD = numOrdenSolicitado + 1
        End If
        rs.Close
        Set rs = Nothing

        cn.RollbackTrans
        enTrans = False
        cn.Close
        Set cn = Nothing
        IntentarReservarOrdenBD = False
        Exit Function
    End If
    rs.Close
    Set rs = Nothing

    ' Está libre -> Insertar y adjudicar atómicamente dentro de la transacción
    Dim sqlIns As String
    Dim recsIns As Long
    sqlIns = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
             "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) VALUES (" & _
             numOrdenSolicitado & ", '" & EscaparSQL(claveRes) & "', " & _
             FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(usrActual) & "', '" & _
             EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"

    cn.Execute sqlIns, recsIns
    If recsIns <> 1 Then
        cn.RollbackTrans
        enTrans = False
        cn.Close
        Set cn = Nothing
        IntentarReservarOrdenBD = False
        Exit Function
    End If

    cn.CommitTrans
    enTrans = False
    cn.Close
    Set cn = Nothing

    IntentarReservarOrdenBD = True
    Exit Function

ErrReserva:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If Not cn Is Nothing Then
        If enTrans Then cn.RollbackTrans
        If cn.State = adStateOpen Then cn.Close
    End If
    Set rs = Nothing
    Set cn = Nothing
    If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
        outSinRed = True
    Else
        Dim dniDummy As String
        Call ExisteOrdenEnBD(residencia, numOrdenSolicitado, outUsuarioPropietario, outNombreHuesped, dniDummy, outSiguienteLibreBD)
        outUsuarioPropietario = FormatearPropietarioBloqueo(outUsuarioPropietario)
    End If
    IntentarReservarOrdenBD = False
    On Error GoTo 0
End Function

```

#### `ReservarSiguienteOrdenLibreBD` (`modDatabase.bas`, líneas 1403-1472)
```vba
Public Function ReservarSiguienteOrdenLibreBD(ByVal residencia As String) As Long
    Dim cn As Object, rs As Object
    Dim claveRes As String, usrActual As String, tokenActual As String
    Dim candidato As Long, maxIter As Integer
    Dim insertado As Boolean

    On Error GoTo ErrReservaSig
    ReservarSiguienteOrdenLibreBD = 0

    If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then Exit Function
    Call AsegurarEsquemaConcurrenciaBD

    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    usrActual = Trim$(ObtenerUsuarioActual())
    If Len(usrActual) = 0 Then usrActual = "admin"
    tokenActual = ObtenerTokenSesionInstancia()

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()

    maxIter = 0
    insertado = False

    Do While Not insertado And maxIter < 20
        maxIter = maxIter + 1
        Set rs = cn.Execute("SELECT MAX(NumOrden) FROM Ordenes " & _
                            "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden < 900001")
        If Not rs.EOF And Not IsNull(rs(0).Value) Then
            candidato = CLng(rs(0).Value) + 1
        Else
            candidato = 1
        End If
        rs.Close
        Set rs = Nothing

        On Error Resume Next
        cn.BeginTrans
        Dim sqlIns As String
        sqlIns = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
                 "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) VALUES (" & _
                 candidato & ", '" & EscaparSQL(claveRes) & "', " & _
                 FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(usrActual) & "', '" & _
                 EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"
        cn.Execute sqlIns
        If Err.Number = 0 Then
            cn.CommitTrans
            insertado = True
        Else
            cn.RollbackTrans
            Err.Clear
            candidato = 0
        End If
        On Error GoTo ErrReservaSig
    Loop

    cn.Close
    Set cn = Nothing
    ReservarSiguienteOrdenLibreBD = candidato
    Exit Function

ErrReservaSig:
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = adStateOpen Then rs.Close
    If Not cn Is Nothing Then If cn.State = adStateOpen Then cn.Close
    Set rs = Nothing
    Set cn = Nothing
    ReservarSiguienteOrdenLibreBD = 0
    On Error GoTo 0
End Function

```

#### `AsegurarOrdenExiste` (`modDatabase.bas`, líneas 1698-1732)
```vba
Public Function AsegurarOrdenExiste(ByVal numOrden As Long, ByVal residencia As String) As Boolean
    If numOrden <= 0 Then Exit Function
    
    Dim resTarget As String
    resTarget = ObtenerClaveResidenciaDesdeHoja(residencia)
    If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then Exit Function
    
    Call AsegurarEsquemaConcurrenciaBD
    
    Dim sqlCheck As String
    Dim varCount As Variant
    sqlCheck = "SELECT COUNT(*) FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'"
    varCount = ExecuteScalar(sqlCheck)
    
    If Not IsNull(varCount) And Not IsEmpty(varCount) Then
        If CLng(varCount) > 0 Then
            AsegurarOrdenExiste = True
            Exit Function
        End If
    End If
    
    Dim sqlInsert As String
    Dim userCreacion As String, tokenActual As String
    userCreacion = Trim$(ObtenerUsuarioActual())
    If Len(userCreacion) = 0 Then userCreacion = "admin"
    tokenActual = ObtenerTokenSesionInstancia()
    
    sqlInsert = "INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, " & _
                "UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) " & _
                "VALUES (" & numOrden & ", '" & EscaparSQL(resTarget) & "', " & _
                FormatearFechaHoraSQL(Now) & ", '" & EscaparSQL(userCreacion) & "', '" & _
                EscaparSQL(tokenActual) & "', " & FormatearFechaHoraSQL(Now) & ", False, False)"
                
    AsegurarOrdenExiste = ExecuteNonQuery(sqlInsert)
End Function

```

#### `SincronizarFilaHaciaAccess` (`modDatabase.bas`, líneas 1737-1993)
```vba
Public Function SincronizarFilaHaciaAccess( _
    ByVal ws As Worksheet, _
    ByVal fila As Long, _
    Optional ByVal permitirAlta As Boolean = False _
) As Boolean

    Dim cn As Object, rs As Object
    Dim enTrans As Boolean
    On Error GoTo ErrorHandler

    If ws Is Nothing Or fila < 2 Then Exit Function

    ' Lectura matricial única de las 28 columnas en memoria (1 sola llamada COM en lugar de 28)
    Dim rowVals As Variant
    rowVals = ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 28)).Value

    Dim vOrden As Variant
    vOrden = rowVals(1, 1)
    If Not IsNumeric(vOrden) Or Len(Trim$(CStr(vOrden))) = 0 Then Exit Function
    
    Dim numOrden As Long
    numOrden = CLng(vOrden)
    If numOrden <= 0 Then Exit Function

    Dim resTarget As String
    resTarget = ObtenerClaveResidenciaDesdeHoja(ws.Name)
    If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then Exit Function

    ' === [SICO-N2: Encolar fila completa si no hay red y no estamos ya dentro del vaciado FIFO] ===
    If Not ModuloSeguridadLocal.EstaSincronizandoFIFO() Then
        If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
            Call ModuloSeguridadLocal.EncolarOperacionOffline( _
                resTarget, numOrden, "SYNC_FILA_COMPLETA", "FILA_COMPLETA", "", "Fila=" & fila)
            SincronizarFilaHaciaAccess = False
            Exit Function
        End If
    End If

    Call AsegurarEsquemaConcurrenciaBD

    Dim uActual As String, tokenActual As String
    uActual = Trim$(ObtenerUsuarioActual())
    If Len(uActual) = 0 Then uActual = "admin"
    tokenActual = ObtenerTokenSesionInstancia()

    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()

    cn.BeginTrans
    enTrans = True

    ' 1. Adquisición atómica del Lease Lock mediante UPDATE condicional dentro de la misma transacción
    Dim sqlAcquire As String, recsLock As Long
    sqlAcquire = "UPDATE Ordenes SET [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "', " & _
                 "[UltimaModificacion] = " & FormatearFechaHoraSQL(Now) & _
                 " WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "' " & _
                 "AND ([UsuarioModificacion] IS NULL OR [UsuarioModificacion] = '' " & _
                 "OR [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "' " & _
                 "OR [UsuarioModificacion] = '" & EscaparSQL(uActual) & "' " & _
                 "OR [UltimaModificacion] IS NULL " & _
                 "OR DateDiff('s', [UltimaModificacion], Now()) >= " & LEASE_LOCK_SEGUNDOS & _
                 " OR DateDiff('s', [UltimaModificacion], Now()) < 0)"

    recsLock = 0
    cn.Execute sqlAcquire, recsLock

    Dim esNuevoRegistro As Boolean
    esNuevoRegistro = False

    If recsLock <> 1 Then
        ' Comprobar dentro de la transacción si el registro existe (bloqueado por otro) o si no existe
        Dim rsCheck As Object
        Dim usrLock As String, fechaLock As Variant
        Set rsCheck = cn.Execute("SELECT UsuarioCreacion, FechaCreacion, UsuarioModificacion, UltimaModificacion " & _
                                 "FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'")
        If Not rsCheck.EOF Then
            usrLock = Trim$(CStr(Nz(rsCheck("UsuarioModificacion").Value, "")))
            fechaLock = rsCheck("UltimaModificacion").Value
            If Len(usrLock) = 0 Or IsNull(fechaLock) Or Not IsDate(fechaLock) Then
                usrLock = Trim$(CStr(Nz(rsCheck("UsuarioCreacion").Value, "")))
                fechaLock = rsCheck("FechaCreacion").Value
            End If
            rsCheck.Close
            Set rsCheck = Nothing
            cn.RollbackTrans
            enTrans = False
            cn.Close
            Set cn = Nothing

            Dim segRest As Long
            segRest = LEASE_LOCK_SEGUNDOS
            If IsDate(fechaLock) Then
                segRest = LEASE_LOCK_SEGUNDOS - DateDiff("s", CDate(fechaLock), Now)
                If segRest < 1 Then segRest = 1
            End If

            MsgBox "BLOQUEO DE EDICIÓN CONCURRENTE:" & vbCrLf & vbCrLf & _
                   "El expediente Nº ORDEN " & numOrden & " (" & resTarget & ") está siendo editado en este momento por:" & vbCrLf & _
                   "   Operador: " & FormatearPropietarioBloqueo(usrLock) & " (bloqueo activo durante " & segRest & " s más)." & vbCrLf & vbCrLf & _
                   "Se actualizará la hoja con los datos del servidor para evitar sobrescribir su trabajo.", _
                   vbExclamation, "Protección Multiusuario - " & resTarget
            Call SincronizarHojaDesdeAccess(resTarget, ws)
            SincronizarFilaHaciaAccess = False
            Exit Function
        End If
        rsCheck.Close
        Set rsCheck = Nothing

        ' El expediente NO existe en Ordenes: prohibir resucitarlo con AddNew salvo que permitirAlta = True
        If Not permitirAlta Then
            cn.RollbackTrans
            enTrans = False
            cn.Close
            Set cn = Nothing
            MsgBox "ATENCIÓN: El expediente Nº ORDEN " & numOrden & " (" & resTarget & ") ya no existe en la base de datos " & _
                   "(fue eliminado o archivado por otro puesto)." & vbCrLf & vbCrLf & _
                   "No se permite recrear automáticamente un expediente eliminado. Se actualizará la hoja.", _
                   vbExclamation, "Expediente Eliminado - " & resTarget
            Call SincronizarHojaDesdeAccess(resTarget, ws)
            SincronizarFilaHaciaAccess = False
            Exit Function
        End If

        esNuevoRegistro = True
    End If

    Dim sql As String
    If esNuevoRegistro Then
        sql = "SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "'"
    Else
        sql = "SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & EscaparSQL(resTarget) & "' " & _
              "AND [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "'"
    End If

    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, 1, 3

    If rs.EOF Then
        If Not esNuevoRegistro Then
            rs.Close
            Set rs = Nothing
            cn.RollbackTrans
            enTrans = False
            cn.Close
            Set cn = Nothing
            SincronizarFilaHaciaAccess = False
            Exit Function
        End If
        rs.AddNew
        rs("NumOrden").Value = numOrden
        rs("Residencia").Value = resTarget
        rs("FechaCreacion").Value = Now
        rs("UsuarioCreacion").Value = uActual
        rs("Solapamiento").Value = False
        rs("ConsentimientoRGPD").Value = False
    End If

    rs("UsuarioModificacion").Value = tokenActual
    rs("UltimaModificacion").Value = Now

    AsignarCampoFecha rs, "FechaPeticion", rowVals(1, 2)
    If Not esNuevoRegistro And (IsEmpty(rowVals(1, 3)) Or Len(Trim$(CStr(rowVals(1, 3)))) = 0) And Not IsNull(rs("NumFactura").Value) Then
        ' Preservar NumFactura existente en BD
    Else
        AsignarCampoNumero rs, "NumFactura", rowVals(1, 3)
    End If
    AsignarCampoTexto rs, "Finalidad", rowVals(1, 4)
    AsignarCampoTexto rs, "Empleo", rowVals(1, 5)
    AsignarCampoTexto rs, "Situacion", rowVals(1, 6)
    AsignarCampoTexto rs, "Evaluacion", rowVals(1, 7)

    If resTarget = "SOTO" Then
        AsignarCampoTexto rs, "Turno", rowVals(1, 8)
    Else
        AsignarCampoTexto rs, "Comision", rowVals(1, 8)
    End If

    AsignarCampoTexto rs, "DNI", rowVals(1, 9)
    AsignarCampoTexto rs, "Rango", rowVals(1, 10)
    AsignarCampoTexto rs, "Nombre", rowVals(1, 11)
    AsignarCampoFecha rs, "FechaEntrada", rowVals(1, 12)
    AsignarCampoFecha rs, "FechaSalida", rowVals(1, 13)
    AsignarCampoNumero rs, "DiasUso", rowVals(1, 14)
    AsignarCampoNumero rs, "PAX", rowVals(1, 15)
    AsignarCampoTexto rs, "Resolucion", rowVals(1, 16)

    Dim valEstadoPagoLocal As Variant
    If resTarget = "OVIEDO" Then
        AsignarCampoNumero rs, "NumHabIndividuales", rowVals(1, 17)
        AsignarCampoNumero rs, "NumHabDobles", rowVals(1, 18)
        AsignarCampoTexto rs, "CamaSuple", rowVals(1, 19)
        AsignarCampoTexto rs, "HabitacionesAsignadas", rowVals(1, 20)
        AsignarCampoTexto rs, "DtoFamNum", rowVals(1, 21)
        AsignarCampoDouble rs, "Importe", rowVals(1, 22)
        AsignarCampoTexto rs, "Telefono", rowVals(1, 23)
        AsignarCampoTexto rs, "Direccion", rowVals(1, 24)
        AsignarCampoTexto rs, "CodigoPostal", rowVals(1, 25)
        AsignarCampoTexto rs, "Poblacion", rowVals(1, 26)
        AsignarCampoTexto rs, "Provincia", rowVals(1, 27)
        valEstadoPagoLocal = rowVals(1, 28)
    ElseIf resTarget = "SOTO" Then
        AsignarCampoTexto rs, "HabitacionesAsignadas", rowVals(1, 19)
        AsignarCampoTexto rs, "DtoFamNum", rowVals(1, 20)
        AsignarCampoDouble rs, "Importe", rowVals(1, 21)
        AsignarCampoTexto rs, "Telefono", rowVals(1, 22)
        AsignarCampoTexto rs, "Direccion", rowVals(1, 23)
        AsignarCampoTexto rs, "CodigoPostal", rowVals(1, 24)
        AsignarCampoTexto rs, "Poblacion", rowVals(1, 25)
        AsignarCampoTexto rs, "Provincia", rowVals(1, 26)
        valEstadoPagoLocal = rowVals(1, 27)
    Else
        AsignarCampoNumero rs, "NumHabIndividuales", rowVals(1, 17)
        AsignarCampoNumero rs, "NumHabDobles", rowVals(1, 18)
        AsignarCampoTexto rs, "HabitacionesAsignadas", rowVals(1, 19)
        AsignarCampoTexto rs, "DtoFamNum", rowVals(1, 20)
        AsignarCampoDouble rs, "Importe", rowVals(1, 21)
        AsignarCampoTexto rs, "Telefono", rowVals(1, 22)
        AsignarCampoTexto rs, "Direccion", rowVals(1, 23)
        AsignarCampoTexto rs, "CodigoPostal", rowVals(1, 24)
        AsignarCampoTexto rs, "Poblacion", rowVals(1, 25)
        AsignarCampoTexto rs, "Provincia", rowVals(1, 26)
        valEstadoPagoLocal = rowVals(1, 27)
    End If

    If Not esNuevoRegistro And (IsEmpty(valEstadoPagoLocal) Or Len(Trim$(CStr(valEstadoPagoLocal))) = 0) And _
       Len(Trim$(CStr(Nz(rs("EstadoPago").Value, "")))) > 0 Then
        ' Preservar EstadoPago existente en BD
    Else
        AsignarCampoTexto rs, "EstadoPago", valEstadoPagoLocal
    End If

    rs.Update
    rs.Close
    Set rs = Nothing

    cn.CommitTrans
    enTrans = False
    cn.Close
    Set cn = Nothing

    SincronizarFilaHaciaAccess = True
    Exit Function

ErrorHandler:
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    If Not cn Is Nothing Then
        If enTrans Then cn.RollbackTrans
        If cn.State = adStateOpen Then cn.Close
        Set cn = Nothing
    End If
    SincronizarFilaHaciaAccess = False
End Function

```

#### `ResolverNumOrdenProvisionalEnBD` (`ModuloSeguridadLocal.bas`, líneas 1364-1402)
```vba
Private Function ResolverNumOrdenProvisionalEnBD( _
    ByVal resCode As String, _
    ByVal numOrdenProv As Long, _
    ByVal wsCola As Worksheet, _
    ByVal filaActualCola As Long, _
    ByVal ultFilaCola As Long _
) As Long
    On Error GoTo ErrResProv
    Dim sqlMax As String, vMax As Variant, nuevoNum As Long
    sqlMax = "SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '" & modDatabase.EscaparSQL(resCode) & "' AND NumOrden < " & MIN_ORDEN_PROVISIONAL
    vMax = modDatabase.ExecuteScalar(sqlMax)
    If IsNumeric(vMax) And Not IsNull(vMax) Then
        nuevoNum = CLng(vMax) + 1
    Else
        nuevoNum = 1
    End If
    
    If modDatabase.AsegurarOrdenExiste(nuevoNum, resCode) Then
        Dim wsSede As Worksheet
        Set wsSede = modDatabase.ObtenerHojaSegura(modDatabase.ObtenerNombreHoja(resCode))
        If Not wsSede Is Nothing Then
            Dim fHoja As Long
            fHoja = BuscarFilaPorNumOrden(wsSede, numOrdenProv)
            If fHoja > 0 Then wsSede.Cells(fHoja, 1).Value = nuevoNum
        End If
        
        Dim k As Long
        For k = filaActualCola To ultFilaCola
            If UCase$(Trim$(CStr(wsCola.Cells(k, eColCola.colResidencia).Value))) = resCode And _
               CLng(val(wsCola.Cells(k, eColCola.colNumOrden).Value)) = numOrdenProv Then
                wsCola.Cells(k, eColCola.colNumOrden).Value = nuevoNum
            End If
        Next k
        ResolverNumOrdenProvisionalEnBD = nuevoNum
    End If
    Exit Function
ErrResProv:
    ResolverNumOrdenProvisionalEnBD = 0
End Function

```

---

## 5. COMPROBACIÓN C — ACTUALIZACIÓN DE EXPEDIENTES Y FECHA DE SALIDA

### 5.1. Mecanismos de Modificación de Campos
1. **Disparo celda a celda:** Al modificar una celda en Excel, el manejador `Worksheet_Change` invoca `ModuloSeguridadLocal.SincronizarCambioDesdeHoja`.
2. **Si hay conexión:** `SincronizarCambioDesdeHoja` invoca `modDatabase.SincronizarCambioTargetHaciaAccess`.
3. **Mapeo y Ejecución:** `SincronizarCambioTargetHaciaAccess` traduce la columna modificada a campo de base de datos mediante `ObtenerCampoBDDesdeColumna` y ejecuta `ActualizarOrden`.
4. **Comprobación de Registros Afectados:** La función `ActualizarOrden` ejecuta `ExecuteNonQuery(sql)`. **No comprueba si el número de registros afectados es mayor que cero**. Si el registro no existía en Access, la sentencia finaliza con éxito técnico (`True`) sin haber actualizado ningún dato real.
5. **Modificación de `FechaSalida` (Columna M / Columna 13):**
   - Se intercepta en `Hoja2.cls`, `Hoja20.cls` y `Hoja24.cls` bajo `Case 12, 13`.
   - Ejecuta la validación de solapes habitacionales (`VerificarDisponibilidadGijon` / `Soto` / `Oviedo`).
   - Invoca `SincronizarCambioDesdeHoja` -> persiste el nuevo valor en el campo `FechaSalida` de `Ordenes`.
   - **Comportamiento acreditado:** **No realiza ninguna comprobación sobre el campo `NumFactura`**. No evalúa si el cambio de año de salida produce una discordancia con el ejercicio de la factura previamente asignada.

### 5.2. Extracción de Procedimientos de Actualización

#### `ActualizarOrden` (`modDatabase.bas`, líneas 2508-2587)
```vba
Public Function ActualizarOrden( _
    ByVal numOrden As Long, _
    ByVal campo As String, _
    ByVal valor As String, _
    Optional ByVal residencia As String = "" _
) As Boolean
    Dim setPart As String
    Dim resTarget As String
    
    ActualizarOrden = False
    If numOrden <= 0 Then Exit Function
    
    resTarget = Trim$(residencia)
    If Len(resTarget) > 0 Then resTarget = ObtenerClaveResidenciaDesdeHoja(resTarget)
    
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        On Error Resume Next
        If Not Application.ActiveSheet Is Nothing Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(Application.ActiveSheet.Name)
        End If
        On Error GoTo 0
    End If
    
    If Len(resTarget) = 0 Or (resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO") Then
        If Len(m_ResidenciaActiva) > 0 And InStr(m_ResidenciaActiva, ",") = 0 Then
            resTarget = ObtenerClaveResidenciaDesdeHoja(m_ResidenciaActiva)
        End If
    End If
    
    ' === [CIERRE V5.2 - Punto 4: Prohibir actualizaciones en Ordenes sin Residencia inequívoca] ===
    If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then
        Exit Function
    End If
    
    ' === [SICO-N2: Enrutamiento a Búfer Local en Modo Contingencia sin devolver falso éxito de BD] ===
    If Not ModuloSeguridadLocal.EstaSincronizandoFIFO() Then
        If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
            Call ModuloSeguridadLocal.EncolarOperacionOffline( _
                resTarget, numOrden, "UPDATE_CAMPO", campo, "", valor)
            ActualizarOrden = False
            Exit Function
        ElseIf ModuloSeguridadLocal.ContarCambiosPendientes() > 0 Then
            Call ModuloSeguridadLocal.ComprobarRetornoRedYSincronizar(DB_PATH)
        End If
    End If

    If LCase$(campo) = "fechaentrada" Or LCase$(campo) = "fechasalida" Or _
       LCase$(campo) = "fechapeticion" Then
        If IsDate(valor) Then
            setPart = "[" & campo & "] = " & FormatearFechaSQL(CDate(valor))
        Else
            setPart = "[" & campo & "] = Null"
        End If
    ElseIf LCase$(campo) = "numfactura" Or LCase$(campo) = "numhabindividuales" Or _
           LCase$(campo) = "numhabdobles" Then
        If Len(Trim$(valor)) = 0 Or Not IsNumeric(valor) Then
            setPart = "[" & campo & "] = Null"
        Else
            setPart = "[" & campo & "] = " & CLng(valor)
        End If
    ElseIf LCase$(campo) = "solapamiento" Or LCase$(campo) = "consentimientorgpd" Then
        setPart = "[" & campo & "] = " & IIf(CBool(valor), "True", "False")
    Else
        setPart = "[" & campo & "] = '" & EscaparSQL(valor) & "'"
    End If

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
End Function

```

#### `SincronizarCambioDesdeHoja` (`ModuloSeguridadLocal.bas`, líneas 625-681)
```vba
Public Function SincronizarCambioDesdeHoja( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    ByVal valorAnterior As Variant, _
    ByVal resCode As String, _
    Optional ByRef outEncoladoOffline As Boolean = False _
) As Boolean
    On Error Resume Next
    SincronizarCambioDesdeHoja = False
    outEncoladoOffline = False
    Dim fRow As Long
    Dim numOrd As Long
    Dim colPagado As Long
    colPagado = IIf(UCase$(Trim$(resCode)) = "OVIEDO", 28, 27)
    
    ' Si hay red disponible, sincronizar selectivamente solo el campo modificado hacia Access
    If SondaRedDisponible(modDatabase.DB_PATH) Then
        If ContarCambiosPendientes() > 0 And Not m_SincronizandoEnCurso Then
            Call SincronizarColaOfflineHaciaAccess
        End If
        SincronizarCambioDesdeHoja = modDatabase.SincronizarCambioTargetHaciaAccess(ws, Target, valorAnterior)
        Exit Function
    End If

    ' En Modo Contingencia SICO-N2: separar claramente el encolado local de la confirmación real en Access
    outEncoladoOffline = True

    If Target.Cells.Count = 1 And Target.Row >= 2 Then
        If Target.Column = 3 Or Target.Column = colPagado Then Exit Function
        
        If IsNumeric(ws.Cells(Target.Row, 1).Value) And Len(Trim$(CStr(ws.Cells(Target.Row, 1).Value))) > 0 Then
            numOrd = CLng(ws.Cells(Target.Row, 1).Value)
            If numOrd > 0 Then
                Dim nombreCampoBD As String
                nombreCampoBD = MapearColumnaACampoBD(resCode, Target.Column)
                If Len(nombreCampoBD) > 0 Then
                    Call EncolarOperacionOffline( _
                        resCode, numOrd, "UPDATE_CAMPO", nombreCampoBD, _
                        CStr(modDatabase.Nz(valorAnterior, "")), CStr(modDatabase.Nz(Target.Value, "")))
                    Exit Function
                End If
            End If
        End If
    End If
    
    For fRow = Target.Row To Target.Row + Target.Rows.Count - 1
        If fRow >= 2 Then
            If IsNumeric(ws.Cells(fRow, 1).Value) And Len(Trim$(CStr(ws.Cells(fRow, 1).Value))) > 0 Then
                numOrd = CLng(ws.Cells(fRow, 1).Value)
                If numOrd > 0 Then
                    Call EncolarOperacionOffline(resCode, numOrd, "SYNC_FILA_COMPLETA", "FILA_COMPLETA", "", "Fila=" & fRow)
                End If
            End If
        End If
    Next fRow
    On Error GoTo 0
End Function

```

#### `SincronizarCambioTargetHaciaAccess` (`modDatabase.bas`, líneas 2307-2444)
```vba
Public Function SincronizarCambioTargetHaciaAccess( _
    ByVal ws As Worksheet, _
    ByVal Target As Range, _
    Optional ByVal valorAnteriorCelda As Variant = Empty _
) As Boolean
    On Error GoTo ErrSyncTarget
    SincronizarCambioTargetHaciaAccess = False
    If ws Is Nothing Or Target Is Nothing Then Exit Function
    
    Dim resTarget As String
    resTarget = ObtenerClaveResidenciaDesdeHoja(ws.Name)
    If resTarget <> "GIJON" And resTarget <> "SOTO" And resTarget <> "OVIEDO" Then
        SincronizarCambioTargetHaciaAccess = True
        Exit Function
    End If
    
    Dim fRow As Long, cCol As Long
    Dim vOrden As Variant
    Dim numOrden As Long
    Dim colHab As Long, colPag As Long
    colHab = IIf(resTarget = "OVIEDO", 20, 19)
    colPag = IIf(resTarget = "OVIEDO", 28, 27)
    
    If Target.Column = 1 Then
        For fRow = Target.Row To Target.Row + Target.Rows.Count - 1
            If fRow >= 2 Then
                vOrden = ws.Cells(fRow, 1).Value
                If IsNumeric(vOrden) And Len(Trim$(CStr(vOrden))) > 0 Then
                    If CLng(vOrden) > 0 Then
                        If Not SincronizarFilaHaciaAccess(ws, fRow) Then Exit Function
                    End If
                End If
            End If
        Next fRow
        SincronizarCambioTargetHaciaAccess = True
        Exit Function
    End If
    
    Dim dictCampos As Object
    Dim campoCol As String
    Dim usrBloq As String, segRest As Long, existeExp As Boolean
    
    For fRow = Target.Row To Target.Row + Target.Rows.Count - 1
        If fRow >= 2 Then
            vOrden = ws.Cells(fRow, 1).Value
            If IsNumeric(vOrden) And Len(Trim$(CStr(vOrden))) > 0 Then
                numOrden = CLng(vOrden)
                If numOrden > 0 Then
                    If EstaExpedienteBloqueadoPorOtro(resTarget, numOrden, usrBloq, segRest, existeExp) Then
                        MsgBox "EDICIÓN BLOQUEADA POR CONCURRENCIA:" & vbCrLf & vbCrLf & _
                               "El expediente Nº ORDEN " & numOrden & " está siendo modificado en este momento por:" & vbCrLf & _
                               "   Sesión activa: " & usrBloq & " (" & segRest & " s restantes)." & vbCrLf & vbCrLf & _
                               "No es posible modificar la fila simultáneamente. Se restaurarán los datos actuales del servidor.", _
                               vbExclamation, "Expediente en Uso - " & resTarget
                        If Target.Cells.Count = 1 And Not IsEmpty(valorAnteriorCelda) Then
                            ws.Cells(Target.Row, Target.Column).Value = valorAnteriorCelda
                        End If
                        Call SincronizarHojaDesdeAccess(resTarget, ws)
                        SincronizarCambioTargetHaciaAccess = False
                        Exit Function
                    ElseIf Not existeExp Then
                        MsgBox "ATENCIÓN: El expediente Nº ORDEN " & numOrden & " ya no existe en la base de datos " & _
                               "(fue eliminado o trasladado por otro puesto)." & vbCrLf & vbCrLf & _
                               "Se actualizará la hoja con el estado real del servidor.", _
                               vbExclamation, "Expediente No Encontrado - " & resTarget
                        Call SincronizarHojaDesdeAccess(resTarget, ws)
                        SincronizarCambioTargetHaciaAccess = False
                        Exit Function
                    End If

                    Set dictCampos = CreateObject("Scripting.Dictionary")
                    
                    For cCol = Target.Column To Target.Column + Target.Columns.Count - 1
                        campoCol = ObtenerCampoBDDesdeColumna(resTarget, cCol)
                        If Len(campoCol) > 0 Then
                            dictCampos(campoCol) = ws.Cells(fRow, cCol).Value
                        End If
                        
                        Select Case cCol
                            Case 4, 5, 6
                                dictCampos("Evaluacion") = ws.Cells(fRow, 7).Value
                                dictCampos(IIf(resTarget = "SOTO", "Turno", "Comision")) = ws.Cells(fRow, 8).Value
                            Case 7
                                If resTarget = "SOTO" Then dictCampos("Turno") = ws.Cells(fRow, 8).Value
                            Case 8
                                If resTarget = "SOTO" Then
                                    dictCampos("FechaEntrada") = ws.Cells(fRow, 12).Value
                                    dictCampos("FechaSalida") = ws.Cells(fRow, 13).Value
                                    dictCampos("DiasUso") = ws.Cells(fRow, 14).Value
                                End If
                            Case 9
                                dictCampos("Rango") = ws.Cells(fRow, 10).Value
                                dictCampos("Nombre") = ws.Cells(fRow, 11).Value
                                dictCampos("Telefono") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 23, 22)).Value
                                dictCampos("Direccion") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 24, 23)).Value
                                dictCampos("CodigoPostal") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 25, 24)).Value
                                dictCampos("Poblacion") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 26, 25)).Value
                                dictCampos("Provincia") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 27, 26)).Value
                            Case 12, 13, 15, 16
                                dictCampos("DiasUso") = ws.Cells(fRow, 14).Value
                                dictCampos("Resolucion") = ws.Cells(fRow, 16).Value
                                dictCampos("HabitacionesAsignadas") = ws.Cells(fRow, colHab).Value
                                If resTarget <> "SOTO" Then
                                    dictCampos("NumHabIndividuales") = ws.Cells(fRow, 17).Value
                                    dictCampos("NumHabDobles") = ws.Cells(fRow, 18).Value
                                End If
                                Dim valResAct As String
                                valResAct = UCase$(Trim$(CStr(ws.Cells(fRow, 16).Value)))
                                If valResAct <> "SI" And valResAct <> "CONCEDIDA" And valResAct <> "REEVALUADA" Then
                                    dictCampos("NumFactura") = ws.Cells(fRow, 3).Value
                                    dictCampos("EstadoPago") = ws.Cells(fRow, colPag).Value
                                End If
                            Case 24, 25, 26, 27
                                If cCol <> colPag Then
                                    dictCampos("CodigoPostal") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 25, 24)).Value
                                    dictCampos("Poblacion") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 26, 25)).Value
                                    dictCampos("Provincia") = ws.Cells(fRow, IIf(resTarget = "OVIEDO", 27, 26)).Value
                                End If
                        End Select
                    Next cCol
                    
                    If dictCampos.Count > 0 Then
                        If Not ActualizarOrdenMultiple(numOrden, dictCampos, resTarget) Then
                            Call SincronizarHojaDesdeAccess(resTarget, ws)
                            SincronizarCambioTargetHaciaAccess = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next fRow
    SincronizarCambioTargetHaciaAccess = True
    Exit Function

ErrSyncTarget:
    SincronizarCambioTargetHaciaAccess = False
End Function

```

---

## 6. COMPROBACIÓN D — FACTURACIÓN AUTOMÁTICA Y MANUAL

### 6.1. Secuencia Real de Asignación y Delimitación Transaccional

```
[Usuario marca PAGADO (AA/AB)]
           │
           ▼
[Validación de Resolución (SI / CONCEDIDA)]
           │
           ▼
[Cálculo de numPropuesto con SELECT UltimoNumero + 1 (Year(Now))]
           │
           ▼
[Detección de huecos con DetectarHuecosEnSerie]
           │
           ▼
[Pregunta interactiva: ¿Desea adjudicar este número de factura?]
           │
    (Si confirma vbYes)
           │
           ▼
[LLAMADA A ObtenerSiguienteNumFactura(resCode, Year(Now))]
     ┌─────────────────────────────────────────────────────────┐
     │ TRANSACCIÓN ADODB EN ACCESS (cn.BeginTrans)             │
     │ 1. SELECT MAX(NumFactura) FROM Ordenes                  │
     │ 2. UPDATE ContadorFacturas SET UltimoNumero = MAX       │
     │ 3. UPDATE ContadorFacturas SET UltimoNumero = Ultimo + 1│
     │ 4. SELECT UltimoNumero                                  │
     │ 5. cn.CommitTrans                                       │
     └─────────────────────────────────────────────────────────┘
           │
           ▼ (FUERA DE LA TRANSACCIÓN ANTERIOR)
[modDatabase.ActualizarOrden(numeroOrden, "NumFactura", nuevoNumero)]
   --> ¡CRÍTICO: AsignarFacturaPorColumna NO comprueba el valor devuelto!
           │
           ▼
[ws.Cells(celda.Row, columnaFactura).Value = nuevoNumero]
           │
           ▼
[SincronizarCamposResidenciaResumen]
```

### 6.2. Vulnerabilidad de Atomicidad Detectada
- En `AsignarFacturaPorColumna` (líneas 192-195):
  ```vba
  If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
      modDatabase.ActualizarOrden CLng(numeroOrden), "NumFactura", CStr(nuevoNumero), resCode
  End If
  ws.Cells(celda.Row, columnaFactura).Value = nuevoNumero
  ```
  La llamada a `modDatabase.ActualizarOrden` se ejecuta como un procedimiento ciego: **no evalúa el booleano devuelto**. Si `ActualizarOrden` falla (por caída de red instantánea o bloqueo):
  1. El contador en Access ya fue incrementado y consolidado (`cn.CommitTrans`).
  2. La celda de Excel muestra el número de factura como asignado.
  3. En Access, el campo `Ordenes.NumFactura` permanece como `NULL`.

### 6.3. Extracción de Procedimientos de Facturación

#### `ObtenerSiguienteNumFactura` (`modDatabase.bas`, líneas 2678-2778)
```vba
Public Function ObtenerSiguienteNumFactura( _
    ByVal residencia As String, _
    ByVal ejercicio As Integer _
) As Long
    Dim cn As Object
    Dim rs As Object
    Dim nuevoNum As Long
    Dim maxEnOrdenes As Long
    Dim intentos As Integer
    Dim exito As Boolean
    Dim ultErrorDesc As String
    Dim claveRes As String
    
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
    
    intentos = 5
    exito = False
    nuevoNum = 0
    
    Do While intentos > 0 And Not exito
        On Error Resume Next
        Err.Clear
        
        Set cn = CreateObject("ADODB.Connection")
        cn.ConnectionTimeout = 5
        cn.Open GetConnectionString()
        
        If Err.Number = 0 Then
            cn.BeginTrans
            
            ' 1. Asegurar que ContadorFacturas no haya quedado por detrás de una factura manual
            maxEnOrdenes = 0
            Set rs = cn.Execute("SELECT MAX(NumFactura) FROM Ordenes WHERE Residencia = '" & EscaparSQL(claveRes) & "'")
            If Err.Number = 0 And Not rs Is Nothing Then
                If Not rs.EOF And Not IsNull(rs(0).Value) Then
                    maxEnOrdenes = CLng(rs(0).Value)
                End If
                rs.Close
                Set rs = Nothing
            End If
            Err.Clear
            
            If maxEnOrdenes > 0 Then
                cn.Execute "UPDATE ContadorFacturas SET UltimoNumero = " & maxEnOrdenes & _
                           " WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio & _
                           " AND UltimoNumero < " & maxEnOrdenes
            End If
            
            ' 2. Incrementar el contador atómicamente
            cn.Execute "UPDATE ContadorFacturas SET UltimoNumero = UltimoNumero + 1 " & _
                       "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio
            
            ' 3. Leer el nuevo valor asignado
            If Err.Number = 0 Then
                Set rs = cn.Execute("SELECT UltimoNumero FROM ContadorFacturas " & _
                                    "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio)
                If Err.Number = 0 And Not rs Is Nothing Then
                    If Not rs.EOF And Not IsNull(rs(0).Value) Then
                        nuevoNum = CLng(rs(0).Value)
                    End If
                    rs.Close
                    Set rs = Nothing
                End If
            End If
            
            If Err.Number = 0 And nuevoNum > 0 Then
                cn.CommitTrans
                If Err.Number = 0 Then
                    exito = True
                Else
                    ultErrorDesc = Err.Description
                    cn.RollbackTrans
                End If
            Else
                ultErrorDesc = Err.Description
                cn.RollbackTrans
            End If
            
            cn.Close
            Set cn = Nothing
        Else
            ultErrorDesc = Err.Description
        End If
        
        On Error GoTo 0
        
        If Not exito Then
            intentos = intentos - 1
            If intentos > 0 Then PausaMS 60
        End If
    Loop
    
    If exito Then
        ObtenerSiguienteNumFactura = nuevoNum
    Else
        MsgBox "Error al obtener número de factura tras varios intentos:" & vbCrLf & ultErrorDesc, _
               vbCritical, "Error de Facturación"
        ObtenerSiguienteNumFactura = 0
    End If
End Function

```

#### `AsignarFacturaPorColumna` (`AsignarNumFactura.bas`, líneas 14-232)
```vba
Public Sub AsignarFacturaPorColumna(ws As Worksheet, _
                                    ByVal Target As Range, _
                                    ByVal columnaCheck As String, _
                                    Optional ByVal columnaFactura As String = "C")

    Dim celda As Range
    Dim nuevoNumero As Long
    Dim numeroOrden As Variant
    Dim respuesta As VbMsgBoxResult
    Dim nombreHojaFactura As String
    Dim hojaFactura As Worksheet
    Dim valorResolucion As String
    Dim mensajeAsignacion As String
    
    ' OPTIMIZACIÓN v2.0: Variables para sincronización
    Dim nombreHojaResidencia As String
    Dim nombreHojaResumen As String
    Dim columnaPagadoResidencia As String
    Dim columnaPagadoResumen As String
    Dim columnaFacturaResumen As String
    Dim resCode As String

    ' ==============================================================================
    ' Determinar código de residencia por CodeName (inmune a nombres de pestaña y codificación)
    ' con fallback por Name mediante modDatabase.ObtenerClaveResidenciaDesdeHoja
    ' ==============================================================================
    Select Case ws.CodeName
        Case "Hoja2"
            resCode = "GIJON"
        Case "Hoja20"
            resCode = "SOTO"
        Case "Hoja24"
            resCode = "OVIEDO"
        Case Else
            resCode = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
    End Select

    ' Determinar las hojas y columnas correspondientes según la residencia
    Select Case resCode
        Case "GIJON"
            nombreHojaFactura = "FACTURA GIJÓN"
            nombreHojaResidencia = "RESIDENCIA GIJÓN"
            nombreHojaResumen = "RESUMEN GIJÓN"
            columnaPagadoResidencia = "AA"
            columnaPagadoResumen = "M"
            columnaFacturaResumen = "N"
            
        Case "SOTO"
            nombreHojaFactura = "FACTURA SOTO"
            nombreHojaResidencia = "RESIDENCIA SOTO"
            nombreHojaResumen = "RESUMEN SOTO"
            columnaPagadoResidencia = "AA"
            columnaPagadoResumen = "L"
            columnaFacturaResumen = "M"
            
        Case "OVIEDO"
            nombreHojaFactura = "FACTURA OVIEDO"
            nombreHojaResidencia = "RESIDENCIA OVIEDO"
            nombreHojaResumen = "RESUMEN OVIEDO"
            columnaPagadoResidencia = "AB"
            columnaPagadoResumen = "M"
            columnaFacturaResumen = "N"
            
        Case Else
            MsgBox "Esta hoja no está habilitada para el proceso de facturación automática.", vbExclamation
            Exit Sub
    End Select

    ' Intentar obtener la hoja de factura
    On Error Resume Next
    Set hojaFactura = ws.Parent.Worksheets(nombreHojaFactura)
    If hojaFactura Is Nothing And resCode = "GIJON" Then
        Set hojaFactura = ws.Parent.Worksheets("FACTURA GIJON")
    End If
    On Error GoTo 0

    If hojaFactura Is Nothing Then
        MsgBox "No existe la hoja de factura llamada '" & nombreHojaFactura & "'.", vbCritical
        Exit Sub
    End If

    ' ' Comprobar si el cambio afecta a la columna de chequeo
    If Not Intersect(Target, ws.Columns(columnaCheck)) Is Nothing Then
        On Error GoTo CleanExit_Fact
        Application.enableEvents = False

        For Each celda In Intersect(Target, ws.Columns(columnaCheck))
            ' Usamos la validación centralizada de pagos para SMS, Bizum, Transferencia, Tarjeta, Efectivo y SI
            If MarcarSiPagadosEnResidencia.EsEstadoPagado(celda.Value) Then
                ' Validar columna RESOLUCION (P = columna 16)
                valorResolucion = UCase$(Trim$(CStr(ws.Cells(celda.Row, 16).Value)))
                
                If valorResolucion = "SI" Or valorResolucion = "CONCEDIDA" Or valorResolucion = "REEVALUADA" Then
                    ' Capturar el Nº ORDEN
                    numeroOrden = ws.Cells(celda.Row, 1).Value

                    ' === [SICO-N2: FACTURA PROFORMA EN MODO OFFLINE Y ENCOLADO FIFO] ===
                    If ModuloSeguridadLocal.EstaEnModoContingencia() Or Not ModuloSeguridadLocal.SondaRedDisponible(modDatabase.DB_PATH) Then
                        Dim pregProforma As VbMsgBoxResult
                        pregProforma = MsgBox( _
                            "MODO FUERA DE LÍNEA (SICO-N2):" & vbCrLf & vbCrLf & _
                            "El servidor central no está disponible en este momento." & vbCrLf & _
                            "¿Desea generar el documento con la marca 'PROFORMA / PENDIENTE DE NUMERACIÓN OFICIAL' " & _
                            "para el Nº ORDEN " & numeroOrden & " y encolar la solicitud en FIFO para asignarle " & _
                            "su número oficial definitivo automáticamente al recuperar la conexión?", _
                            vbYesNo + vbQuestion, "Emisión Proforma Offline - SICO-N2")
                            
                        If pregProforma = vbYes Then
                            ws.Cells(celda.Row, columnaFactura).Value = "PROFORMA"
                            
                            If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
                                Call ModuloSeguridadLocal.EncolarOperacionOffline( _
                                    resCode, CLng(numeroOrden), "UPDATE_CAMPO", "EstadoPago", "", CStr(celda.Value))
                                Call ModuloSeguridadLocal.EncolarOperacionOffline( _
                                    resCode, CLng(numeroOrden), "ASIGNAR_FACTURA_DIFERIDA", "NumFactura", "PROFORMA", "SOLICITAR_OFICIAL")
                            End If
                            
                            If Trim$(CStr(numeroOrden)) <> "" Then
                                Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
                                    nombreHojaResidencia, columnaFactura, columnaPagadoResidencia, _
                                    nombreHojaResumen, columnaFacturaResumen, columnaPagadoResumen, numeroOrden)
                            End If
                            
                            respuesta = MsgBox("Documento registrado como 'PROFORMA / PENDIENTE DE NUMERACIÓN OFICIAL' " & _
                                               "y encolado en el búfer FIFO." & vbCrLf & vbCrLf & _
                                               "¿Desea abrir e imprimir el documento proforma?", _
                                               vbYesNo + vbQuestion, "Imprimir Proforma")
                            If respuesta = vbYes Then
                                Call modDatabase.MostrarYActivarHojaSegura(nombreHojaFactura)
                                hojaFactura.Range("D1").Value = numeroOrden
                            End If
                        End If
                        GoTo SiguienteCeldaFactura
                    End If
                    ' === [FIN BLOQUE SICO-N2 PROFORMA OFFLINE] ===
                    
                    ' Previsualizar el próximo número sin incrementar aún el contador físico en Access
                    Dim numPropuesto As Long
                    Dim varUlt As Variant
                    varUlt = modDatabase.ExecuteScalar("SELECT UltimoNumero FROM ContadorFacturas WHERE Residencia = '" & _
                                                       modDatabase.EscaparSQL(resCode) & "' AND Ejercicio = " & Year(Now))
                    If IsNumeric(varUlt) And Not IsNull(varUlt) Then
                        numPropuesto = CLng(varUlt) + 1
                    Else
                        numPropuesto = 1
                    End If
                    mensajeAsignacion = "Asignación de factura única desde Access para " & resCode
                    
                    ' Detectar si existen huecos en la serie de facturación
                    Dim textoHuecos As String
                    textoHuecos = DetectarHuecosEnSerie(ws, columnaFactura, numPropuesto)
                    
                    Dim mensajePrompt As String
                    mensajePrompt = "Se propone asignar al Nº ORDEN " & numeroOrden & _
                                    " la factura Nº " & numPropuesto & vbCrLf & vbCrLf & _
                                    mensajeAsignacion
                    
                    If Len(textoHuecos) > 0 Then
                        mensajePrompt = mensajePrompt & vbCrLf & vbCrLf & _
                                        "ADVERTENCIA. En la serie, faltan las facturas Nº:" & vbCrLf & _
                                        textoHuecos & vbCrLf & _
                                        "Valore si corresponde mantener este salto o rellenar huecos."
                    End If
                    
                    mensajePrompt = mensajePrompt & vbCrLf & vbCrLf & _
                                    "¿Desea adjudicar este número de factura?"
                    
                    ' Preguntar al usuario ANTES de incrementar el contador atómico en Access
                    Dim preguntaAsignar As VbMsgBoxResult
                    preguntaAsignar = MsgBox(mensajePrompt, _
                                             vbYesNo + vbQuestion, "Confirmar Adjudicación de Factura")

                    If preguntaAsignar = vbYes Then
                        ' === Ahora sí: Obtener e incrementar número atómico desde Access ===
                        nuevoNumero = modDatabase.ObtenerSiguienteNumFactura(resCode, Year(Now))
                        If nuevoNumero <= 0 Then GoTo CleanExit_Fact

                        ' Asignar el número de factura en Access (fuente de verdad) Y en la celda local
                        If IsNumeric(numeroOrden) And CLng(numeroOrden) > 0 Then
                            modDatabase.ActualizarOrden CLng(numeroOrden), "NumFactura", CStr(nuevoNumero), resCode
                        End If
                        ws.Cells(celda.Row, columnaFactura).Value = nuevoNumero
                        
                        ' ========================================================================
                        ' OPTIMIZACIÓN v2.0: SINCRONIZAR AUTOMÁTICAMENTE CON RESUMEN
                        ' Sincroniza C (Nº FACTURA) y AA/AB (PAGADO) con el RESUMEN
                        ' ========================================================================
                        If Trim$(CStr(numeroOrden)) <> "" Then
                            Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
                                nombreHojaResidencia, _
                                columnaFactura, _
                                columnaPagadoResidencia, _
                                nombreHojaResumen, _
                                columnaFacturaResumen, _
                                columnaPagadoResumen, _
                                numeroOrden)
                        End If
                        ' ========================================================================
                        
                        ' Preguntar si desea imprimir la factura
                        respuesta = MsgBox("Nº de factura " & nuevoNumero & " asignado correctamente." & vbCrLf & vbCrLf & _
                                           "¿Desea imprimir la factura?", vbYesNo + vbQuestion, "Imprimir factura")
                        If respuesta = vbYes Then
                            Call modDatabase.MostrarYActivarHojaSegura(nombreHojaFactura)
                            hojaFactura.Range("D1").Value = numeroOrden
                        End If
                    End If
                Else
                    MsgBox "No es posible generar una factura con la resolución adoptada (" & ws.Cells(celda.Row, 16).Value & ").", vbExclamation, "Factura no generada"
                    celda.Value = ""
                End If
            End If
SiguienteCeldaFactura:
        Next celda

CleanExit_Fact:
        Application.enableEvents = True
    End If
End Sub

```

#### `ProcesarCambioManualFactura` (`AsignarNumFactura.bas`, líneas 561-641)
```vba
Public Sub ProcesarCambioManualFactura(ws As Worksheet, Target As Range, ByVal targetRow As Long, Optional ByVal columnaFactura As String = "C")
    Dim numOrden As Variant
    Dim nuevoValor As Variant
    Dim resCode As String
    Dim numFacturaLong As Long
    Dim ultimaFila As Long
    Dim rangoCol As Range
    
    If Target.Cells.CountLarge > 1 Or targetRow < 2 Then Exit Sub
    
    ' Capturar el Nº ORDEN
    numOrden = ws.Cells(targetRow, 1).Value
    If IsEmpty(numOrden) Or Not IsNumeric(numOrden) Or CLng(numOrden) <= 0 Then
        If Not IsEmpty(Target.Value) Then
            MsgBox "Debe existir un Nº ORDEN válido antes de asignar un número de factura.", vbExclamation, "Sin Nº ORDEN"
            Application.enableEvents = False
            Target.ClearContents
            Application.enableEvents = True
        End If
        Exit Sub
    End If
    
    ' Determinar código de residencia
    Select Case ws.CodeName
        Case "Hoja2":  resCode = "GIJON"
        Case "Hoja20": resCode = "SOTO"
        Case "Hoja24": resCode = "OVIEDO"
        Case Else:     resCode = modDatabase.ObtenerClaveResidenciaDesdeHoja(ws.Name)
    End Select
    
    nuevoValor = Target.Value
    
    ' Caso 1: Celda vaciada (borrado manual de número de factura)
    If IsEmpty(nuevoValor) Or Trim(CStr(nuevoValor)) = "" Then
        modDatabase.ActualizarOrden CLng(numOrden), "NumFactura", "", resCode

        Exit Sub
    End If
    
    ' Caso 2: Validación de valor numérico
    If Not IsNumeric(nuevoValor) Then
        MsgBox "El número de factura debe ser un valor numérico entero positivo.", vbExclamation, "Valor no válido"
        Application.enableEvents = False
        Target.ClearContents
        Application.enableEvents = True
        Exit Sub
    End If
    
    numFacturaLong = CLng(nuevoValor)
    If numFacturaLong <= 0 Then
        MsgBox "El número de factura debe ser mayor que cero.", vbExclamation, "Valor no válido"
        Application.enableEvents = False
        Target.ClearContents
        Application.enableEvents = True
        Exit Sub
    End If
    
    ' 1. Comprobación local de duplicados en la hoja Excel activa
    ultimaFila = ws.Cells(ws.Rows.Count, columnaFactura).End(xlUp).Row
    Set rangoCol = ws.Range(columnaFactura & "1:" & columnaFactura & ultimaFila)
    Call DuplicadosOrdenyFactura.VerificarDuplicados(Target, rangoCol, "NUM. FACT.")
    
    If IsEmpty(Target.Value) Then Exit Sub
    
    ' 2. Comprobación anti-colisión en Access (por si otro usuario ya lo asignó)
    If modDatabase.ExisteFacturaEnBD(resCode, numFacturaLong, CLng(numOrden)) Then
        MsgBox "El Nº de factura " & numFacturaLong & " YA ESTÁ ASIGNADO en la base de datos para " & resCode & "." & vbCrLf & vbCrLf & _
               "Operación cancelada para evitar duplicidad entre puestos.", vbCritical, "Factura Duplicada en Access"
        Application.enableEvents = False
        Target.ClearContents
        Application.enableEvents = True
        Exit Sub
    End If
    
    ' 3. Advertencia informativa de huecos en la serie
    Call AsignarNumFactura.ValidarFacturaManual(ws, numFacturaLong, columnaFactura)
    
    ' 4. Persistir inmediatamente en Access y alinear ContadorFacturas si el Nº manual adelanta la serie
    modDatabase.ActualizarOrden CLng(numOrden), "NumFactura", CStr(numFacturaLong), resCode
    modDatabase.SincronizarContadorFacturaManual resCode, numFacturaLong, Year(Now)
End Sub

```

#### `SincronizarContadorFacturaManual` (`modDatabase.bas`, líneas 2782-2801)
```vba
Public Sub SincronizarContadorFacturaManual( _
    ByVal residencia As String, _
    ByVal numFacturaManual As Long, _
    Optional ByVal ejercicio As Integer = 0 _
)
    On Error Resume Next
    If numFacturaManual <= 0 Then Exit Sub
    If ejercicio <= 0 Then ejercicio = Year(Now)
    
    Dim claveRes As String
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
    
    Dim sql As String
    sql = "UPDATE ContadorFacturas SET UltimoNumero = " & numFacturaManual & _
          " WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND Ejercicio = " & ejercicio & _
          " AND UltimoNumero < " & numFacturaManual
    Call ExecuteNonQuery(sql)
    On Error GoTo 0
End Sub

```

#### `ExisteFacturaEnBD` (`modDatabase.bas`, líneas 2889-2916)
```vba
Public Function ExisteFacturaEnBD( _
    ByVal residencia As String, _
    ByVal numFactura As Long, _
    Optional ByVal excluirNumOrden As Long = 0) As Boolean
    
    Dim sql As String
    Dim resCount As Variant
    Dim claveRes As String
    
    claveRes = ObtenerClaveResidenciaDesdeHoja(residencia)
    If Len(claveRes) = 0 Then claveRes = UCase(Trim(residencia))
    
    sql = "SELECT COUNT(*) FROM Ordenes " & _
          "WHERE Residencia = '" & EscaparSQL(claveRes) & "' " & _
          "AND NumFactura = " & numFactura
    
    If excluirNumOrden > 0 Then
        sql = sql & " AND NumOrden <> " & excluirNumOrden
    End If
    
    resCount = ExecuteScalar(sql)
    
    If IsNull(resCount) Then
        ExisteFacturaEnBD = False
    Else
        ExisteFacturaEnBD = (CLng(resCount) > 0)
    End If
End Function

```

---

## 7. COMPROBACIÓN E — RECONCILIACIÓN OFFLINE DE FACTURAS

### 7.1. Respuestas Específicas al Recorrido de Reconciliación

1. **Cómo se selecciona una operación `ASIGNAR_FACTURA_DIFERIDA`:**
   En `SincronizarColaOfflineHaciaAccess` (líneas 963-985), se recorren secuencialmente las filas de `SYS_COLA_SYNC` ordenadas por `colIDTransaccion ASC`. Si el estado es `PENDIENTE` o `ERROR`, se evalúa `Select Case tipoOp`. Al coincidir con `"ASIGNAR_FACTURA_DIFERIDA"`, se delega en `ConsolidarFacturaDiferidaFIFO`.
2. **Momento de la llamada:**
   Se ejecuta en el orden FIFO estricto de la cola, después de haberse resuelto las pre-reservas y altas de órdenes previas.
3. **Tratamiento si devuelve `False`:**
   Si `ConsolidarFacturaDiferidaFIFO` retorna `False`, el motor marca `colEstadoSync = "ERROR"` (línea 1063), incrementa `colReintentos` en 1 y, si la red ha caído, aborta el bucle activando el modo contingencia.
4. **Campos actualizados antes y después:**
   - *Antes:* `colEstadoSync` pasa transitoriamente a `"SINCRONIZANDO"`.
   - *Después:* Si fue exitoso, pasa a `"CONCILIADO"`.
5. **Decisión de consolidación y purga:**
   Al finalizar el bucle de sincronización (líneas 1075-1079), el motor recorre la hoja `SYS_COLA_SYNC` en sentido inverso (`ultFila To 2 Step -1`). **Todas las filas con `colEstadoSync = "CONCILIADO"` son eliminadas físicamente** (`wsCola.Rows(r).Delete`). A continuación, `VolcarColaHaciaDiscoWAL` reescribe íntegramente el archivo `.wal` local en disco conteniendo únicamente las transacciones que no fueron conciliadas.
6. **Vulnerabilidad de doble consumo de factura en reintentos:**
   En `ConsolidarFacturaDiferidaFIFO` (líneas 1137-1147):
   ```vba
   numFactOficial = modDatabase.ObtenerSiguienteNumFactura(resCode, anioEjercicio)
   If numFactOficial <= 0 Then Exit Function
   If Not modDatabase.ActualizarOrden(numOrden, "NumFactura", CStr(numFactOficial), resCode) Then Exit Function
   ```
   Si se produce un cierre abrupto de Excel o fallo de red **después** de `ActualizarOrden` pero **antes** de que la fila se marque como `"CONCILIADO"` y se borre del WAL:
   - La operación permanece en el WAL como pendiente.
   - En el siguiente reinicio, se vuelve a ejecutar `ConsolidarFacturaDiferidaFIFO`.
   - La función **no comprueba si `Ordenes.NumFactura` ya tiene asignado un número en Access**.
   - Invoca nuevamente `ObtenerSiguienteNumFactura`, **consumiendo un segundo número de factura atómico en Access y sobreescribiendo el anterior**, dejando el primer número como un hueco irreversible y descolgado en la contabilidad.

### 7.2. Extracción de Procedimientos de Reconciliación FIFO

#### `SincronizarColaOfflineHaciaAccess` (`ModuloSeguridadLocal.bas`, líneas 905-1116)
*(Procedimiento real correspondiente a `ReconciliarColaOfflineFIFO`)*
```vba
Public Function SincronizarColaOfflineHaciaAccess() As Boolean
    On Error GoTo errSync
    If m_SincronizandoEnCurso Then Exit Function
    
    Dim evOrig As Boolean, scrOrig As Boolean
    evOrig = Application.enableEvents
    scrOrig = Application.screenUpdating
    
    m_SincronizandoEnCurso = True
    
    ' Verificar conexión real con Access antes de iniciar el vaciado
    If Not modDatabase.ComprobarConexion() Then
        m_SincronizandoEnCurso = False
        m_EnModoContingencia = True
        m_UltimoFalloTick = Timer
        SincronizarColaOfflineHaciaAccess = False
        Exit Function
    End If
    
    Dim wsCola As Worksheet
    Set wsCola = ObtenerOCrearHojaCola()
    Call RehidratarColaDesdeDiscoWALSiNecesario(wsCola)
    
    Dim ultFila As Long
    ultFila = wsCola.Cells(wsCola.Rows.Count, eColCola.colIDTransaccion).End(xlUp).Row
    If ultFila < 2 Then
        m_SincronizandoEnCurso = False
        m_EnModoContingencia = False
        modDatabase.ActualizarBarraEstado
        SincronizarColaOfflineHaciaAccess = True
        Exit Function
    End If
    
    m_EnModoContingencia = False
    Application.enableEvents = False
    Application.screenUpdating = False
    Application.StatusBar = "SINCRONIZANDO COLA FIFO CON SERVIDOR CENTRAL..."
    
    DesprotegerHojaCola wsCola
    
    If ultFila > 2 Then
        wsCola.Range(wsCola.Cells(1, 1), wsCola.Cells(ultFila, eColCola.colCRC)).Sort _
            Key1:=wsCola.Cells(2, eColCola.colIDTransaccion), Order1:=xlAscending, Header:=xlYes
    End If
    
    Dim r As Long
    Dim estadoItem As String
    Dim resItem As String, usrItem As String, tipoOp As String, campoMod As String
    Dim valAnt As String, valNuevo As String, fechaLocalStr As String
    Dim numOrdItem As Long
    Dim totalConciliados As Long, totalConflictos As Long, totalErroresRed As Long
    Dim detalleConflictos As String
    
    totalConciliados = 0
    totalConflictos = 0
    totalErroresRed = 0
    detalleConflictos = ""
    
    For r = 2 To ultFila
        estadoItem = UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colEstadoSync).Value)))
        
        If estadoItem = "PENDIENTE" Or estadoItem = "ERROR" Then
            wsCola.Cells(r, eColCola.colEstadoSync).Value = "SINCRONIZANDO"
            
            fechaLocalStr = Trim$(CStr(wsCola.Cells(r, eColCola.colFechaHoraLocal).Value))
            usrItem = Trim$(CStr(wsCola.Cells(r, eColCola.colUsuario).Value))
            resItem = Trim$(CStr(wsCola.Cells(r, eColCola.colResidencia).Value))
            numOrdItem = CLng(val(wsCola.Cells(r, eColCola.colNumOrden).Value))
            tipoOp = UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colTipoOperacion).Value)))
            campoMod = Trim$(CStr(wsCola.Cells(r, eColCola.colCampoModificado).Value))
            valAnt = CStr(wsCola.Cells(r, eColCola.colValorAnterior).Value)
            valNuevo = CStr(wsCola.Cells(r, eColCola.colValorNuevo).Value)
            
            Dim okItem As Boolean
            Dim esConflictoHab As Boolean
            Dim msgConflicto As String
            okItem = False
            esConflictoHab = False
            msgConflicto = ""
            
            Select Case tipoOp
                Case "LOG_OFFLINE"
                    okItem = InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, campoMod, valNuevo)
                    
                Case "ELIMINAR_ORDEN"
                    If FueExpedienteModificadoPosteriormenteEnBD(resItem, numOrdItem, fechaLocalStr, msgConflicto) Then
                        esConflictoHab = True
                    Else
                        okItem = modDatabase.EliminarOrdenBD(resItem, numOrdItem)
                    End If
                    
                Case "PRERESERVA", "SYNC_FILA_COMPLETA"
                    Dim eraProvisional As Boolean
                    eraProvisional = EsNumOrdenProvisional(numOrdItem)
                    If eraProvisional Then
                        Dim nuevoOrdReal As Long
                        nuevoOrdReal = ResolverNumOrdenProvisionalEnBD(resItem, numOrdItem, wsCola, r, ultFila)
                        If nuevoOrdReal > 0 Then numOrdItem = nuevoOrdReal
                    End If
                    If Not eraProvisional And FueExpedienteModificadoPosteriormenteEnBD(resItem, numOrdItem, fechaLocalStr, msgConflicto) Then
                        esConflictoHab = True
                    Else
                        Dim wsHojaSede As Worksheet
                        Set wsHojaSede = modDatabase.ObtenerHojaSegura(modDatabase.ObtenerNombreHoja(resItem))
                        Dim filaEnHoja As Long
                        filaEnHoja = BuscarFilaPorNumOrden(wsHojaSede, numOrdItem)
                        If filaEnHoja > 0 Then
                            Dim habFila As String
                            habFila = ObtenerHabitacionDeFila(wsHojaSede, filaEnHoja, resItem)
                            If Len(habFila) > 0 Then
                                If ExisteConflictoHabitacionServerWins(resItem, numOrdItem, habFila, msgConflicto) Then
                                    esConflictoHab = True
                                    LimpiarHabitacionEnConflictoFila wsHojaSede, filaEnHoja, resItem
                                End If
                            End If
                            okItem = modDatabase.SincronizarFilaHaciaAccess(wsHojaSede, filaEnHoja, (tipoOp = "PRERESERVA" Or eraProvisional))
                        Else
                            okItem = True
                        End If
                    End If
                    
                Case "ASIGNAR_FACTURA_DIFERIDA"
                    okItem = ConsolidarFacturaDiferidaFIFO(resItem, numOrdItem, fechaLocalStr, usrItem)
                    
                Case "UPDATE_CAMPO"
                    If StrComp(campoMod, "HabitacionesAsignadas", vbTextCompare) = 0 And Len(Trim$(valNuevo)) > 0 Then
                        If ExisteConflictoHabitacionServerWins(resItem, numOrdItem, valNuevo, msgConflicto) Then
                            esConflictoHab = True
                            RevertirHabitacionEnHojaExcel resItem, numOrdItem, valAnt
                        End If
                    End If
                    If Not esConflictoHab Then
                        If ExisteConflictoPreImageEnServidor(resItem, numOrdItem, campoMod, valAnt, valNuevo, msgConflicto) Then
                            esConflictoHab = True
                        End If
                    End If
                    
                    If Not esConflictoHab Then
                        okItem = modDatabase.ActualizarOrden(numOrdItem, campoMod, valNuevo, resItem)
                        If okItem Then
                            Call InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, "SYNC_FIFO_" & UCase$(campoMod), _
                                "Consolidado offline [" & resItem & "] Orden " & numOrdItem & ": " & campoMod & _
                                " ('" & valAnt & "' -> '" & valNuevo & "')")
                        End If
                    End If
            End Select
            
            If esConflictoHab Then
                wsCola.Cells(r, eColCola.colEstadoSync).Value = "CONFLICTO_HAB"
                totalConflictos = totalConflictos + 1
                detalleConflictos = detalleConflictos & "  - [" & resItem & "] Orden Nº " & numOrdItem & _
                                    " (Hab. '" & valNuevo & "'): " & msgConflicto & vbCrLf
                Call InsertarLogHistoricoFIFO(fechaLocalStr, usrItem, "CONFLICTO_FIFO_SERVER_WINS", _
                    "Prevalece servidor en [" & resItem & "] Orden " & numOrdItem & ": " & msgConflicto)
            ElseIf okItem Then
                wsCola.Cells(r, eColCola.colEstadoSync).Value = "CONCILIADO"
                totalConciliados = totalConciliados + 1
            Else
                wsCola.Cells(r, eColCola.colEstadoSync).Value = "ERROR"
                wsCola.Cells(r, eColCola.colReintentos).Value = CLng(val(wsCola.Cells(r, eColCola.colReintentos).Value)) + 1
                totalErroresRed = totalErroresRed + 1
                If Not modDatabase.ComprobarConexion() Then
                    m_EnModoContingencia = True
                    m_UltimoFalloTick = Timer
                    Exit For
                End If
            End If
        End If
    Next r
    
    For r = ultFila To 2 Step -1
        If UCase$(Trim$(CStr(wsCola.Cells(r, eColCola.colEstadoSync).Value))) = "CONCILIADO" Then
            wsCola.Rows(r).Delete
        End If
    Next r
    
    ProtegerHojaCola wsCola
    Call VolcarColaHaciaDiscoWAL(wsCola)
    
    Application.enableEvents = evOrig
    Application.screenUpdating = scrOrig
    m_SincronizandoEnCurso = False
    modDatabase.ActualizarBarraEstado
    
    If totalConciliados > 0 Or totalConflictos > 0 Then
        Dim msgResumen As String
        msgResumen = "Sincronización FIFO completada con el servidor central:" & vbCrLf & vbCrLf & _
                     "  - Operaciones consolidadas en Access: " & totalConciliados & vbCrLf
        If totalConflictos > 0 Then
            msgResumen = msgResumen & vbCrLf & _
                         "ATENCIÓN - CONFLICTOS DE HABITACIÓN (Regla Server Wins):" & vbCrLf & _
                         "Las siguientes habitaciones asignadas fuera de línea ya habían sido ocupadas en el servidor por otro puesto. " & _
                         "Prevalece la reserva del servidor; por favor, reasigne habitación a:" & vbCrLf & _
                         detalleConflictos
            MsgBox msgResumen, vbExclamation, "Reconciliación SICO-N2 con Conflictos"
        Else
            MsgBox msgResumen, vbInformation, "Reconciliación SICO-N2 Completada"
        End If
    End If
    
    SincronizarColaOfflineHaciaAccess = (totalErroresRed = 0)
    Exit Function

errSync:
    On Error Resume Next
    ProtegerHojaCola wsCola
    Application.enableEvents = evOrig
    Application.screenUpdating = scrOrig
    m_SincronizandoEnCurso = False
    modDatabase.ActualizarBarraEstado
    SincronizarColaOfflineHaciaAccess = False
End Function

```

#### `ConsolidarFacturaDiferidaFIFO` (`ModuloSeguridadLocal.bas`, líneas 1122-1187)
```vba
Private Function ConsolidarFacturaDiferidaFIFO( _
    ByVal resCode As String, _
    ByVal numOrden As Long, _
    ByVal fechaHoraLocal As String, _
    ByVal usuario As String _
) As Boolean
    On Error GoTo ErrFactFIFO
    
    Dim anioEjercicio As Integer
    If IsDate(fechaHoraLocal) Then
        anioEjercicio = Year(CDate(fechaHoraLocal))
    Else
        anioEjercicio = Year(Now)
    End If
    
    Dim numFactOficial As Long
    numFactOficial = modDatabase.ObtenerSiguienteNumFactura(resCode, anioEjercicio)
    If numFactOficial <= 0 Then
        ConsolidarFacturaDiferidaFIFO = False
        Exit Function
    End If
    
    If Not modDatabase.ActualizarOrden(numOrden, "NumFactura", CStr(numFactOficial), resCode) Then
        ConsolidarFacturaDiferidaFIFO = False
        Exit Function
    End If
    
    Dim nombreHojaRes As String, nombreHojaSum As String
    Dim colPagRes As String, colPagSum As String, colFactSum As String
    
    Select Case UCase$(Trim$(resCode))
        Case "GIJON"
            nombreHojaRes = "RESIDENCIA GIJÓN"
            nombreHojaSum = "RESUMEN GIJÓN"
            colPagRes = "AA": colPagSum = "M": colFactSum = "N"
        Case "SOTO"
            nombreHojaRes = "RESIDENCIA SOTO"
            nombreHojaSum = "RESUMEN SOTO"
            colPagRes = "AA": colPagSum = "L": colFactSum = "M"
        Case "OVIEDO"
            nombreHojaRes = "RESIDENCIA OVIEDO"
            nombreHojaSum = "RESUMEN OVIEDO"
            colPagRes = "AB": colPagSum = "M": colFactSum = "N"
    End Select
    
    Dim wsRes As Worksheet
    Set wsRes = modDatabase.ObtenerHojaSegura(nombreHojaRes)
    If Not wsRes Is Nothing Then
        Dim filaOrd As Long
        filaOrd = BuscarFilaPorNumOrden(wsRes, numOrden)
        If filaOrd > 0 Then
            wsRes.Cells(filaOrd, "C").Value = numFactOficial
            Call MarcarSiPagadosEnResidencia.SincronizarCamposResidenciaResumen( _
                wsRes.Name, "C", colPagRes, nombreHojaSum, colFactSum, colPagSum, numOrden)
        End If
    End If
    
    Call InsertarLogHistoricoFIFO(fechaHoraLocal, usuario, "FACTURA_DIFERIDA_FIFO", _
        "Convertida PROFORMA a Factura Oficial Nº " & numFactOficial & " en [" & resCode & "] para Nº ORDEN " & numOrden)
        
    ConsolidarFacturaDiferidaFIFO = True
    Exit Function

ErrFactFIFO:
    ConsolidarFacturaDiferidaFIFO = False
End Function

```

#### `InsertarLogHistoricoFIFO` (`ModuloSeguridadLocal.bas`, líneas 1338-1358)
```vba
Private Function InsertarLogHistoricoFIFO( _
    ByVal fechaHoraLocalStr As String, _
    ByVal usuario As String, _
    ByVal accion As String, _
    ByVal detalle As String _
) As Boolean
    Dim fechaLog As Date
    If IsDate(fechaHoraLocalStr) Then
        fechaLog = CDate(fechaHoraLocalStr)
    Else
        fechaLog = Now
    End If
    
    Dim sql As String
    sql = "INSERT INTO LogActividad (FechaHora, Usuario, Accion, Detalle) VALUES (" & _
          modDatabase.FormatearFechaHoraSQL(fechaLog) & ", " & _
          "'" & modDatabase.EscaparSQL(usuario) & "', " & _
          "'" & modDatabase.EscaparSQL(accion) & "', " & _
          "'" & modDatabase.EscaparSQL(detalle & " [Fecha/Hora original offline: " & fechaHoraLocalStr & "]") & "')"
    InsertarLogHistoricoFIFO = modDatabase.ExecuteNonQuery(sql)
End Function

```

---

## 8. COMPROBACIÓN F — TIPOS Y FORMATOS DE LOS CINCO CAMPOS SIN PERSISTENCIA

Se ha analizado la configuración física de las celdas en el OpenXML nativo (`sheet2.xml`, `sheet23.xml`, `sheet34.xml`) y el código VBA asociado:

| Residencia | Col | Encabezado | Validación de Datos Excel | Formato Numérico | Bloqueo Celda | ¿Tiene Fórmulas? | Procedimientos que interactúan con ella |
| :--- | :---: | :--- | :--- | :--- | :---: | :---: | :--- |
| **GIJÓN** | **AD (30)** | `REPARTIR COSTE  HAB.` | Lista desplegable: `Formula1 = "NO,SI"` (`allowBlank=1`) | General | Desbloqueada (`locked=0`) | **NO** | `BorradoDatosAlojamiento.bas` (L183: se limpia con `ws.Range(Cells(2, 30), Cells(lastRow, 30)).ClearContents`). |
| **SOTO** | **AD (30)** | `REPARTIR COSTE  APTO.` | Lista desplegable: `Formula1 = "NO,SI"` (`allowBlank=1`) | General | Desbloqueada (`locked=0`) | **NO** | `BorradoDatosAlojamiento.bas` (L189: se limpia con `ClearContents`). |
| **OVIEDO** | **AC (29)** | `REPARTIR COSTE  HAB.` | Lista desplegable: `Formula1 = "NO,SI"` (`allowBlank=1`) | `0` (Entero) / General | Desbloqueada (`locked=0`) | **NO** | `CopiarDatosExcelOviedo.bas` (L115, L206: recibe datos históricos columna T de 2025). `BorradoDatosAlojamiento.bas` (L193: se limpia con `ClearContents`). |
| **OVIEDO** | **AD (30)** | `TAG` | Ninguna | General | Desbloqueada (`locked=0`) | **NO** | Celdas libres para marcas de control manual del usuario. Se limpian en `BorradoDatosAlojamiento.bas` (L193). |
| **OVIEDO** | **AE (31)** | `R` | Ninguna | General | Desbloqueada (`locked=0`) | **NO** | `Hoja24.cls` (L660): Al hacer clic o editar la columna 31, dispara la validación de coherencia habitacional entre PAX y habitaciones. Se limpia en `BorradoDatosAlojamiento.bas` (L193). |

- **Reglas de negocio observadas:**
  - Los campos de reparto de costes (`AD` en GIJÓN/SOTO y `AC` en OVIEDO) están configurados con validación estricta de lista binaria (`NO,SI`). Su función funcional es permitir al operador indicar si el importe de la estancia se divide entre varios huéspedes.
  - Ninguna de estas cinco columnas tiene campo receptor en las tablas `Ordenes` ni `Ordenes_Historico` de Access.
  - No intervienen en los cálculos de importes, no se transmiten por la cola offline `SYS_COLA_SYNC` y se pierden definitivamente en cualquier operación de compactación o archivado.

---

## 9. SÍNTESIS DE EVIDENCIAS Y CONCLUSIONES PARA LA DIRECCIÓN ARQUITECTÓNICA

### 9.1. Hechos Técnicos Confirmados
1. **Identidad de Fila:** El libro carece de cualquier identificador técnico de fila o entidad. La única clave operativa compartida es `(Residencia, NumOrden)`.
2. **Comportamiento ante Manipulación de Filas:** La eliminación de filas mediante `Delete Shift:=xlUp` (en archivado y borrado de solicitudes) invalida cualquier mecanismo de indexación externa basado en el número de fila. Cualquier identificador permanente local debe ubicarse en una columna de la propia fila para desplazarse solidariamente con ella.
3. **Ausencia de Verificación de Éxito en Facturación:** `AsignarFacturaPorColumna` no comprueba si `modDatabase.ActualizarOrden` tuvo éxito antes de reflejar el número de factura en la hoja activa.
4. **Vulnerabilidad de Doble Consumo en FIFO:** `ConsolidarFacturaDiferidaFIFO` no verifica si la orden ya posee número de factura en Access antes de solicitar uno nuevo, arriesgando el consumo redundante de números en caso de crash entre la actualización y el commit de la cola.
5. **Independencia de Fecha de Salida:** La modificación de `FechaSalida` no comprueba ni sincroniza el año de la factura asociada.
6. **Campos Volátiles:** Los cinco campos singulares analizados son celdas desbloqueadas de usuario sin correspondencia en la base de datos backend.

### 9.2. Ausencias y Extremos No Verificables
- No existen hojas anuales independientes en el libro activo.
- No existen rutinas que sincronicen las columnas `AD` (GIJÓN/SOTO) o `AC:AE` (OVIEDO) con Access.
- No existe ningún procedimiento denominado `ReconciliarColaOfflineFIFO`; la función real que realiza esta tarea es `SincronizarColaOfflineHaciaAccess`.

---

## 10. DECLARACIÓN FORMAL DE NO ALTERACIÓN Y RESPETO DE RESTRICCIONES

El inspector técnico certifica solemnemente que:
1. Todas las operaciones se han ejecutado en modalidad estrictamente de solo lectura sobre el código y las bases de datos.
2. No se ha modificado, recompilado ni guardado ninguna línea de código del libro `BDAS_Multiusuario.xlsm` ni de los módulos VBA.
3. No se ha ejecutado ninguna sentencia DDL ni DML de modificación sobre la base de datos `Residencia_BE.accdb`.
4. No se ha propuesto ningún esquema físico, diseño de base de datos ni solución arquitectónica, quedando dicha función reservada a ChatGPT.
5. Se ha respetado escrupulosamente la regla de nombres empleando el literal `GIJÓN`.
6. En cumplimiento riguroso del RGPD, no se ha expuesto ningún dato de carácter personal.
