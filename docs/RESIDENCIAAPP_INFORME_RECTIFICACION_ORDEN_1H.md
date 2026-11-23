# RESIDENCIAAPP — INFORME DE RECTIFICACIÓN ARQUITECTÓNICA DIRECTA
## Orden 1H — Revisión de identidad del módulo y rectificación del Documento Maestro v0.2.7

**Fecha:** 16 de noviembre de 2026  
**Autoría de la rectificación documental:** Dirección Arquitectónica (ChatGPT)  
**Revisión complementaria:** Orden 1J (precisiones documentales sobre CodeName y EstadoPago)  
**Documento base:** `DOCUMENTO_MAESTRO_ARQUITECTURA_REQUISITOS_RESIDENCIAAPP_v0.2.6.md` (contenido íntegro extraído del archivo compuesto aportado en la conversación)  
**Documento resultante:** `RESIDENCIAAPP_DOCUMENTO_MAESTRO_v0.2.7.md`  
**Alcance:** rectificación documental, sin modificar VBA, Excel, Access, binarios de producción, Version Reference, Golden Master ni repositorio.

---

## 1. Dictamen ejecutivo

La v0.2.6 no podía mantenerse como documento de referencia definitivo porque el bloque de persistencia no describía correctamente las rutinas integradas en el Front-End de producción. La evidencia recibida el 16/11/2026 cierra la identidad del módulo y permite resolver las discrepancias principales sin pedir una nueva generación documental a Antigravity.

**Dictamen:** se ha preparado la v0.2.7 como candidata a revisión arquitectónica independiente. No se declara aprobada, homologada ni certificada funcionalmente. Esta rectificación es documental y no incluye ejecución de pruebas de Excel/Access.

## 2. Evidencia de identidad del módulo

El informe recibido acredita:

| Elemento | Valor |
|---|---|
| Front-End activo | `BDAS_Multiusuario.xlsm` |
| SHA-256 del Front-End | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` |
| `modDatabase.bas` del repositorio | 298.794 bytes; 7.231 líneas |
| `modDatabase` extraído del binario | 298.794 bytes; 7.231 líneas |
| SHA-256 de ambos módulos | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` |
| Cotejo byte a byte | Coincidencia exacta; 0 diferencias |

Esta identidad convierte el módulo extraído en fuente As-Built fiable para el contenido inspeccionado en la Orden 1H.

## 3. Correspondencia física y persistencia de las tres sedes

Se incorporó el mapeo A:AB de las rutinas de persistencia y se mantuvo la información previa de columnas técnicas fuera de ese rango.

- Las tres hojas tienen encabezados verificados en el intervalo A:AC. Esto **no** afirma que todas terminen en AC: el Documento Maestro conserva la evidencia previa de columnas AD:AG en OVIEDO.
- Ambas rutinas (`SincronizarFilaHaciaAccess` y `SincronizarLoteImportacionBD`) leen las 28 primeras columnas A:AB.
- La columna C (`NÚM FACT.`) se asigna a `NumFactura`.
- `EstadoPago` procede de AA en GIJÓN/SOTO y AB en OVIEDO.
- En SOTO Q y R se leen en memoria, pero se omiten deliberadamente en el mapeo a Access.
- AC y las columnas posteriores quedan fuera de las matrices leídas.

La correspondencia definitiva de hojas, CodeName VBA y partes XML del libro activo queda fijada así:

| Nombre visible | CodeName VBA | Parte XML |
|---|---|---|
| RESIDENCIA GIJÓN | `Hoja2` | `xl/worksheets/sheet2.xml` |
| RESIDENCIA SOTO | `Hoja20` | `xl/worksheets/sheet23.xml` |
| RESIDENCIA OVIEDO | `Hoja24` | `xl/worksheets/sheet34.xml` |

Además, `Hoja1` corresponde a Calendario GIJÓN, `Hoja23` a Calendario OVIEDO y `Hoja34` a TARIFAS UNIDAD. La extracción de la Orden 1H confundió nombres de partes XML con CodeName VBA; el volcado estructural del libro permite separarlos. No se requiere renombrar módulos ni modificar el libro.

La matriz del Documento Maestro distingue la persistencia individual de la persistencia por lotes y mantiene la variación de posiciones/campos de OVIEDO.

## 4. Rectificación crítica: `SincronizarLoteImportacionBD`

La extracción integrada acredita que la rutina:

1. Abre una conexión y ejecuta `cn.BeginTrans`.
2. Abre un Recordset `SELECT * FROM Ordenes WHERE Residencia = ...`.
3. Indexa los registros existentes por `NumOrden`.
4. Ejecuta `AddNew` para los no existentes y `rs.Update` para cada fila procesada.
5. Confirma el conjunto con `cn.CommitTrans`; ante error, ejecuta `cn.RollbackTrans`.
6. No utiliza `rs.UpdateBatch` en el código integrado aportado.
7. No adquiere ni valida el Lease Lock por expediente. En registros existentes, escribe el token actual en `UsuarioModificacion` y renueva `UltimaModificacion`.

**Conclusión:** la transacción protege el lote frente a errores de ejecución, pero no sustituye el control lógico de concurrencia por expediente. Una importación masiva puede actualizar una orden cuyo Lease Lock está vigente en otra instancia. Se añade `DEF-FIFO-07`.

Esta es la discrepancia más importante de la revisión: una transacción de base de datos no debe documentarse como si fuera un Lease Lock.

## 5. Rectificación de factura y estado de pago

Se mantiene la constatación de que no se han acreditado procedimientos con los nombres `ActualizarFacturaBD` o `ActualizarEstadoPagoBD`. Pero la conclusión anterior de que no existía persistencia separada para esos campos era incorrecta.

Las dos rutinas verificadas asignan:

- `NumFactura`: columna C, con preservación del valor en Access cuando la celda llega vacía para una fila existente.
- `EstadoPago`: columna AA en GIJÓN/SOTO y AB en OVIEDO, también con preservación del valor previo cuando la celda llega vacía para una fila existente.

En los registros nuevos, la rutina por lotes no fija por sí misma el valor predeterminado `PENDIENTE`; asigna el contenido de la celda mediante `AsignarCampoTexto`.

Esto no significa que ninguna otra ruta del sistema inicialice ese estado: `IntentarReservarOrdenBD`, cuando el número solicitado está libre y sigue la rama `INSERT`, escribe expresamente `EstadoPago = 'PENDIENTE'` (líneas 1358–1369 de `modDatabase.bas`). Esta evidencia se limita a esa ruta de reserva y no demuestra que Access tenga configurado un `DefaultValue` general en el campo. Cuando `forzarReserva=True` y el número ya existe, la rutina sigue una rama `UPDATE` distinta.

## 6. Lease Lock individual: precondiciones y ramas

La sentencia de producción compara `UsuarioModificacion` con NULL, cadena vacía y el token de instancia `tokenMod`; **no** contiene un predicado `UsuarioModificacion = uActual` ni un fallback `admin` en ese `WHERE`. También contempla `UltimaModificacion IS NULL`, caducidad de 120 segundos y marca temporal futura (`DateDiff < 0`).

Para afirmar que el `UPDATE` afecta cero filas por bloqueo vigente deben cumplirse simultáneamente estas precondiciones: `UsuarioModificacion` no nulo/no vacío y distinto de `tokenMod`; `UltimaModificacion` no nula; y `DateDiff('s', UltimaModificacion, Now())` entre 0 y 119.

Se precisa además que:

- `UsuarioModificacion` nulo/vacío o igual al token de instancia permite adquisición; también la permite `UltimaModificacion IS NULL`.
- `DateDiff < 0` es una condición exclusiva de producción para marcas temporales futuras; no aparece en el SQL V5.2 de laboratorio.
- La firma real de `SincronizarFilaHaciaAccess` solo recibe `ws`, `fila` y `resTarget`; no tiene parámetro `permitirAlta`.
- Si `recsLock <> 1`, la función ejecuta un SELECT posterior, pero no distingue `EOF`: muestra el mensaje genérico de conflicto, revierte, llama a `SincronizarOrdenHaciaExcel(ws, numOrden, fila, resTarget)` y retorna `False`. El registro inexistente no activa una rama de `AddNew` en ese punto.
- El `AddNew` que existe en la función aparece después de una adquisición satisfactoria y de abrir el Recordset de la clave. `AsegurarOrdenExiste` es un helper separado; no se le atribuye invocación automática desde esta función sin evidencia del llamador.

## 7. CapKey, Certeza y homologación

Se corrige `REQ-OBJ-03`: «Aprobado en Laboratorio» se sustituye por «Aprobado en Diseño; código presente en laboratorio; homologación funcional NO ACREDITADA (Certeza 4)». Se conserva la taxonomía de certeza y no se convierte la inspección estática de código en certificación de pruebas.

## 8. Limitaciones y cuestiones que permanecen abiertas

La revisión complementaria 1J ha resuelto la discrepancia de CodeName; no se mantiene como incidencia abierta. Permanecen estas limitaciones:

1. **Alcance de la extracción:** la evidencia de la Orden 1H transcribe las dos rutinas de persistencia y verifica el hash del módulo completo. No vuelve a transcribir todos los restantes procedimientos de `modDatabase.bas`; por tanto, sus descripciones heredadas no se presentan como revalidadas por la Orden 1H.
2. **Git de Orden 1F:** sigue siendo **NO ACREDITADO**. No se reconstruye retrospectivamente una captura inicial inexistente.
3. **Homologación V5.2/CapKey:** permanece en Certeza 4. No se aportan registros de pruebas reproducibles nuevos que justifiquen ascenderla a Certeza 2.
4. **Valor predeterminado de Access:** se acredita la asignación explícita de `PENDIENTE` en la rama de inserción de `IntentarReservarOrdenBD`, pero no se ha acreditado en esta revisión un `DefaultValue` general configurado en el esquema de Access.

## 9. Integridad y alcance de los entregables

Se ha extraído el texto completo del Documento Maestro v0.2.6 desde el adjunto compuesto aportado en la conversación, se han aplicado las rectificaciones de la Orden 1H y las precisiones complementarias de la Orden 1J, y se han mantenido dos archivos Markdown independientes. La versión 0.2.6 se conserva como antecedente; no se declara sobrescrita en el repositorio del proyecto.

**No se ha realizado:** edición de código VBA, apertura/guardado del libro Excel, modificación de Access, pruebas funcionales, compilación COM, cambios en Git, modificación del Golden Master ni de Version Reference.

---

## 10. Veredicto

**ORDEN 1H — RECTIFICACIÓN DOCUMENTAL DIRECTA COMPLETADA; PRECISIONES COMPLEMENTARIAS 1J INCORPORADAS.**

La v0.2.7 corrige la identidad y el mapeo del módulo de persistencia, la cobertura de la rutina por lotes, el tratamiento de `NumFactura`/`EstadoPago`, el alcance real del Lease Lock y la clasificación de CapKey. La Orden 1J resuelve la correspondencia de CodeName y precisa que `IntentarReservarOrdenBD` inicializa `EstadoPago` con `PENDIENTE` en su rama de inserción de una reserva libre. Se mantiene `DEF-FIFO-07` por ausencia de protección Lease Lock en importación masiva.

**Estado de la v0.2.7: CANDIDATA A REVISIÓN ARQUITECTÓNICA INDEPENDIENTE.** No se declara aprobada ni homologada; permanecen las limitaciones indicadas en el apartado 8.


---

## 11. Rectificación complementaria — Orden 1J

La Orden 1J introduce exclusivamente estas dos precisiones respecto del documento emitido en la Orden 1H:

1. **CodeName:** se cierra la correspondencia de las hojas maestras como `Hoja2` (RESIDENCIA GIJÓN), `Hoja20` (RESIDENCIA SOTO) y `Hoja24` (RESIDENCIA OVIEDO). Se identifica el error de la extracción 1H como confusión entre el nombre de la parte XML (`sheetN.xml`) y el CodeName VBA. No se modifica ni renombra ningún componente.
2. **EstadoPago:** se mantiene que la rutina por lotes no establece `PENDIENTE` por sí misma; se añade que `IntentarReservarOrdenBD` sí lo escribe expresamente en la rama `INSERT` de un número libre. No se extrapola esta evidencia a un valor predeterminado general del campo Access.

No se modifica el resto del contenido técnico de la v0.2.7, el recuento de 18 defectos, `DEF-FIFO-07`, el Anexo A ni la clasificación Certeza 4 de CapKey/V5.2. Los hashes de esta revisión complementaria sustituyen a los de los archivos originales de la Orden 1H para futuras incorporaciones; los originales quedan preservados en el paquete ZIP de entrega de la Orden 1H y en la copia histórica previa.


## 12. Rectificación complementaria — Orden 1K

Durante la revisión independiente posterior a la incorporación de la Orden 1J se detectó una discrepancia puntual en el recuento de reintentos de `ReservarSiguienteOrdenLibreBD`: el Documento Maestro indicaba hasta 20 iteraciones, mientras que el código integrado en producción establece `maxIntentos = 5` en la línea 1425 y limita el bucle a cinco iteraciones (líneas 1429–1467).

Se corrigen únicamente las dos descripciones As-Built afectadas: apartado 2.7 y ficha `DEF-ORD-01`. Se conserva el diagnóstico de `MAX+1` y la ausencia de búsqueda de huecos. La arquitectura objetivo ORD-10, el catálogo de 18 defectos, `DEF-FIFO-07`, la clasificación Certeza 4 y la condición de candidata de la v0.2.7 no cambian.

**Estado:** precisión documental incorporada; no se han realizado modificaciones de código ni pruebas funcionales.


## 13. Rectificación complementaria — Orden 1K y Orden 1L

### 13.1. Conciliación de `ConfigTarifasCeldas` (Orden 1K)

La consulta de solo lectura ejecutada en producción obtuvo 291 registros: 155 en `TARIFAS UNIDAD` y 136 en `TURNOS-PRECIOS`. El recuento se concilia con los rangos efectivos del código:

- `TARIFAS UNIDAD`: `A3:C7` (15) + `E3:N9` (70) + `E13:N19` (70) = 155. Las columnas `K:N` de GIJÓN y `M:N` de OVIEDO forman parte de los rangos reales; las 42 coordenadas no son externas a la configuración efectiva.
- `TURNOS-PRECIOS`: `A1:F20` (120 posiciones, 17 fórmulas excluidas = 103 registros) + `A22:F23` (12) + `B25:D31` (21) = 136.

Total: 155 + 136 = 291. No se leyeron campos con los importes o valores de tarifas ni se realizaron escrituras en Access.

### 13.2. Rectificación del Lease Lock de producción (Orden 1L)

Se corrigen las afirmaciones del apartado 6 de este informe que atribuían al SQL de producción un fallback `uActual` y un parámetro `permitirAlta`. La extracción literal de `SincronizarFilaHaciaAccess` demuestra que:

1. El `WHERE` solo compara `UsuarioModificacion` con NULL, cadena vacía y el token de instancia; no incorpora `UsuarioModificacion = uActual`.
2. La función no tiene parámetro `permitirAlta`.
3. Cuando `recsLock <> 1`, el SELECT posterior no discrimina `EOF`; la función muestra el mismo mensaje genérico, revierte, llama a `SincronizarOrdenHaciaExcel` y retorna `False`.
4. El `AddNew` posterior no es una recuperación de ese rechazo; aparece después de que el `UPDATE` de adquisición haya afectado una fila.
5. `AsegurarOrdenExiste` es independiente y no se presupone su llamada desde `SincronizarFilaHaciaAccess` sin evidencia del llamador.

Se mantiene la diferenciación entre la condición de marca temporal futura (`DateDiff < 0`) en producción y su ausencia en el SQL de laboratorio V5.2.

**Estado documental:** los apartados 3.4, 3.7 y la fila 8 del Anexo A de la v0.2.7 quedan rectificados. El documento sigue siendo candidato a revisión arquitectónica independiente. No se modifica VBA, Excel, Access ni Git.
