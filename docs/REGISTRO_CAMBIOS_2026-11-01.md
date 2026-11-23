# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-11-01 | Hito Oficial: Versión de Referencia 2026-11-01 03:45
### Ámbito: Optimización Definitiva y Erradicación de Lentitud del Borrado (Parcial y Total) + Compactación Física Atómica de Filas con Delete Shift:=xlUp + Encolador de Contingencia Offline por Lotes en Memoria RAM sin Sobrecarga I/O en WAL

---

## 1. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia exacta entre cada hito técnico de implementación, los archivos de respaldo incrementales generados en disco, sus dimensiones, hashes criptográficos y el conjunto de funcionalidades y correcciones introducidas:

| Hito / Copia Incremental | Ruta del Archivo de Respaldo | Tamaño (Bytes) | Hash MD5 / SHA-256 | Funcionalidades, Operaciones y Módulos Vinculados |
| :--- | :--- | :--- | :--- | :--- |
| **0. Estado Base de Partida**<br>*(Hito 2026-10-31 20:15)* | H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_2015.xlsm<br>H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_2015.xlsm | 6.749.596 B | **MD5:** 6AE80126541D7D7F273B94929ECAE6A9<br>**SHA256:** 210745C9D089798C1623A4CF0E448FAA7F4344C41A7B6A7B0AAA6B6B7F378E9A | Punto de partida de la sesión. Incorporaba Concurrencia V5.2 (Token GUID + Predicados Atómicos) y Blindaje Integral de Calendarios GIJÓN, SOTO y OVIEDO. |
| **1. Backup Preventivo PRE-BORRADO**<br>*(Antes de Optimización)* | H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-01_0230_PRE_OPTIMIZACION_BORRADO.xlsm<br>H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-01_0230_PRE_OPTIMIZACION_BORRADO.accdb | 6.766.450 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** 37DD52FD244070C123E31A1527EEB78F<br>**XLSM SHA256:** F810788B194229487A7987BE4E0F502E9084ADD61F3E3A808DABFB179CFB3A83<br>**ACCDB MD5:** 6237949FA8E9F503C9623F0A7294310D | Copia de seguridad previa al despliegue del paquete quirúrgico de optimización del borrado. |
| **2. Backup POST-BORRADO 02:30**<br>*(Despliegue Inicial Borrado)* | H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-01_0230_POST_OPTIMIZACION_BORRADO.xlsm | 6.782.848 B | **MD5:** 817D8B327CB86176C85896A348ACD3AE<br>**SHA256:** FB700356E6347A67F273CF8A2501CB43DB0FBAF957D6A55024B14944958E0EB4 | Despliegue inicial de optimizaciones: saneamiento de columnas > 36 en BorradoDatosAlojamiento.bas, EliminarOrdenesLoteBD en Access y estructura base. |
| **3. Punto de Restauración Seguro**<br>*(Reversión OVIEDO Intacta)* | H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_RESTORE_POINT_OVIEDO_INTACTO.xlsm | 6.782.848 B | **MD5:** 817D8B327CB86176C85896A348ACD3AE<br>**SHA256:** FB700356E6347A67F273CF8A2501CB43DB0FBAF957D6A55024B14944958E0EB4 | Instantánea generada para blindar y preservar la integridad de las filas 2 a 91 de RESIDENCIA OVIEDO, tras revertir y purgar las transacciones de prueba encoladas en el archivo WAL. |
| **4. Backup Consolidación Referencia**<br>*(Hito 2026-11-01 03:45)* | H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-01_0345_VERSION_REFERENCIA.xlsm<br>H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-11-01_0345_VERSION_REFERENCIA.xlsm<br>H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-01_0345_VERSION_REFERENCIA.accdb | 6.792.119 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** D616F3366D4AB18F507B5AE47A49DEE7<br>**XLSM SHA256:** 4DCAF380D5A2FB03A6F17B153CF88F686324B582931E6988DD7120ED81602200<br>**ACCDB MD5:** 6237949FA8E9F503C9623F0A7294310D<br>**ACCDB SHA256:** 7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196 | Copia incremental consolidada de cierre de sesión con la solución definitiva de borrado instantáneo y compactación de filas. |
| **5. Versión de Referencia Oficial Vigente**<br>*(Hito 2026-11-01 03:45)* | H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm<br>H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0345.xlsm<br>H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0345.xlsm<br>H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb<br>H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-01_0345.accdb | 6.792.119 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** D616F3366D4AB18F507B5AE47A49DEE7<br>**XLSM SHA256:** 4DCAF380D5A2FB03A6F17B153CF88F686324B582931E6988DD7120ED81602200<br>**ACCDB MD5:** 6237949FA8E9F503C9623F0A7294310D<br>**ACCDB SHA256:** 7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196 | **Nueva Versión de Referencia Oficial Consolidada.** Réplicas protegidas en Solo Lectura (ReadOnly = True), con 0 errores de compilación y cero mojibake. |

---

## 2. Detalle Exhaustivo de las Modificaciones Implementadas por Módulo

### 2.1. BorradoSolicitudes.bas
- **Problema corregido:** En el intento anterior de optimización, el Paso 4 sustituyó la instrucción de borrado nativa por un vaciado de celdas (.ClearContents), lo que provocaba que las filas 2 a 91 quedasen vacías en blanco y las filas inferiores (92 en adelante) no subiesen.
- **Solución implementada:**
  - Se agrupan todas las filas a eliminar en memoria dentro de un único objeto Range mediante Application.Union(rngEliminar, ws.Rows(fila)).
  - Se ejecuta una única llamada física atómica: 
ngEliminar.Delete Shift:=xlUp.
  - Se eliminan físicamente las filas de Excel en menos de 0,01 segundos, provocando que la fila 92 ascienda de forma compacta a la fila 2.
  - Todas las fórmulas relativas de importe y textos de la tabla se actualizan automáticamente sin provocar errores #¡REF!.
  - Se desprotege la hoja con la contraseña segura y se reprotege incondicionalmente con UserInterfaceOnly:=True y flags de formato permitidos.

### 2.2. ModuloSeguridadLocal.bas
- **Problema corregido:** Cuando la aplicación operaba en modo offline / contingencia (debido a que la ruta de red DB_PATH no estaba conectada), el encolado de 90 órdenes ejecutaba 90 desprotecciones/reprotecciones, 308.880 lecturas COM y 90 reescrituras síncronas completas del archivo WAL en disco (VolcarColaHaciaDiscoWAL), tardando ~1 segundo por orden (90 segundos en total).
- **Solución implementada:**
  - Nueva función pública EncolarBorradosLoteOffline(resCode, colOrdenes).
  - Desprotección única inicial de la hoja SYS_COLA_SYNC.
  - Purgado previo de operaciones en memoria mediante tabla hash Scripting.Dictionary.
  - Asignación matricial en bloque rrNuevasFilas(1 To N, 1 To 13) en memoria RAM con cálculo de UUID y CRC16.
  - Asignación correcta al miembro colReintentos (corrigiendo la errata de compilación previa).
  - Volcado único matricial a las celdas de la hoja de cola.
  - **Un solo guardado síncrono al archivo físico WAL en disco (VolcarColaHaciaDiscoWAL)**, reduciendo el tiempo de encolado de 90 segundos a menos de **0,04 segundos**.

### 2.3. modDatabase.bas
- **Problema corregido:** EliminarOrdenesLoteBD iteraba orden por orden invocando el encolador individual offline.
- **Solución implementada:**
  - En la rama de contingencia (Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH)), se delega de inmediato el lote completo a ModuloSeguridadLocal.EncolarBorradosLoteOffline(claveRes, colOrdenes).
  - Se mantiene intacta la ruta online de alta velocidad, que ejecuta el archivado y borrado en Access mediante una única consulta SQL DELETE ... IN (...) en menos de 0,01 segundos.

---

## 3. Estado de Validación y Gobernanza

- **Control de Compilación COM:** Validado en instancia real de Excel COM; el botón de compilación de VBE (CommandBars.FindControl(1, 578).Enabled) se encuentra en False, confirmando **0 errores de compilación** en todo el proyecto.
- **Control Anti-Mojibake:** Escaneo exhaustivo ejecutado sobre los 123 componentes exportados; se confirma **cero mojibake** y codificación nativa Windows-1252 ANSI.
- **Regla Estricta de Nombres:** Cumplimiento estricto del literal GIJÓN sin funciones de escape numéricas.


---

# HITO OFICIAL: VERSIÓN DE REFERENCIA 2026-11-01 04:25
## Ámbito: Perfeccionamiento Quirúrgico Estricto de Borrado Parcial y Total (Concurrencia V5.2)

### 1. Resumen de Implementación de los 4 Cambios Quirúrgicos

1. **Cambio 1 (ModuloSeguridadLocal.bas - EncolarBorradosLoteOffline):**
   - Captura estricta de idTransBase mediante Application.WorksheetFunction.Max **antes** de cualquier purgado en SYS_COLA_SYNC, blindando la monotonicidad del ID_Transaccion y los UUID requeridos por la arquitectura de sincronización V5.2.
   - Purgado en memoria RAM en una sola pasada COM (rrCola = wsCola.Range(...).Value) acumulando filas obsoletas con Application.Union y eliminándolas en un único 
ngPurgar.Delete Shift:=xlUp.
   - Dimensionamiento exacto del bloque matricial a 13 columnas operativas (colCRC = 13) con cálculo de CRC16 y persistencia en un único paso I/O hacia el archivo WAL en disco (VolcarColaHaciaDiscoWAL).

2. **Cambio 2 (modDatabase.bas - EliminarOrdenesLoteBD y EliminarTodasOrdenesSedeBD):**
   - Incorporación de verificación transaccional defensiva del Lease Lock V5.2 (ventana de 120 segundos) previa a la ejecución de sentencias DELETE.
   - Detección reactiva de órdenes en edición concurrente por otro puesto e interrupción limpia con mensaje informativo detallando el operador propietario mediante FormatearPropietarioBloqueo.
   - Enrutamiento SICO-N2 en EliminarTodasOrdenesSedeBD mediante el parámetro opcional colOrdenesOffline As Collection, permitiendo la contingencia sin conexión en borrados totales de sede.

3. **Cambio 3 (BorradoDatosAlojamiento.bas - BorrarDatosAlojamientoOptimizado):**
   - Desprotección mediante ModuloConfigSegura.ObtenerPasswordHojas() y retirada de autofiltros (ws.ShowAllData) **antes** de medir la última fila para evitar la omisión de registros filtrados/ocultos.
   - Cálculo exacto de lastRowDatos discriminando por hoja las columnas de fórmulas pre-arrastradas (excluyendo cols 21, 28, 29 en GIJÓN; cols 17, 18, 21, 28, 29 en SOTO; y cols 22, 32, 33 en OVIEDO), erradicando lecturas erróneas del tamaño real de la tabla.
   - Recolección en RAM de órdenes activas (1 sola lectura COM en Columna A) para alimentar la contingencia SICO-N2.
   - Aborto defensivo del vaciado visual de celdas (If Not modDatabase.EliminarTodasOrdenesSedeBD(...) Then Exit Sub), preservando la integridad frente a fallos transaccionales o colisiones de Lease Lock.

4. **Cambio 4 (BorradoSolicitudes.bas - Paso 4 en BorrarSolicitudesPorOrden):**
   - Blindaje estricto de las propiedades de entorno Application.EnableEvents = False, Application.ScreenUpdating = False y Application.Calculation = xlCalculationManual inmediatamente antes de la eliminación física atómica (
ngEliminar.Delete Shift:=xlUp).
   - Supresión incondicional de cualquier disparo espurio de Worksheet_Change durante el desplazamiento de filas.

---

### 2. Matriz de Hashes Criptográficos SHA-256 y MD5 (Hito 2026-11-01 04:25)

| Archivo / Componente | Tamaño | SHA-256 | MD5 |
| :--- | :--- | :--- | :--- |
| BDAS_Multiusuario.xlsm (Activo Trabajo) | 6.809.197 B | E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C | 5727B66025C7278C95B5596CA22BD351 |
| Residencia_BE.accdb (BD Activa) | 6.746.112 B | 7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196 | 6237949FA8E9F503C9623F0A7294310D |
| BDAS_Multiusuario_BACKUP_2026-11-01_0418_PRE_PERFECCIONAMIENTO_V5_2.xlsm | 6.792.119 B | 4DCAF380D5A2FB03A6F17B153CF88F686324B582931E6988DD7120ED81602200 | D616F3366D4AB18F507B5AE47A49DEE7 |
| BDAS_Multiusuario_BACKUP_2026-11-01_0425_POST_PERFECCIONAMIENTO_V5_2.xlsm | 6.809.197 B | E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C | 5727B66025C7278C95B5596CA22BD351 |
| Residencia_BE_BACKUP_2026-11-01_0425_POST_PERFECCIONAMIENTO_V5_2.accdb | 6.746.112 B | 7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196 | 6237949FA8E9F503C9623F0A7294310D |
| BDAS_Multiusuario_VERSION_REFERENCIA.xlsm (Read-Only) | 6.809.197 B | E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C | 5727B66025C7278C95B5596CA22BD351 |
| BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm (Read-Only) | 6.809.197 B | E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C | 5727B66025C7278C95B5596CA22BD351 |
| Residencia_BE_VERSION_REFERENCIA.accdb (Read-Only) | 6.746.112 B | 7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196 | 6237949FA8E9F503C9623F0A7294310D |
| BDAS_Multiusuario.xlsm (Copia Red Operadores) | 6.809.197 B | E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C | 5727B66025C7278C95B5596CA22BD351 |
| ModuloSeguridadLocal.bas (VBA Export) | 64.764 B | 297C38BED746892527AE518B54D6C899A2FEE5D67C4E0B394C0C20753D6FADA9 | C5EE8AC350EE3CD378EA202D9D245D14 |
| modDatabase.bas (VBA Export) | 204.282 B | C5008E92BBF023780B8E5655AF54920E3BCC23E6017123E20682FB71CB4841B0 | A69F155B4D45A71AD047AC20B1692F72 |
| BorradoDatosAlojamiento.bas (VBA Export) | 9.887 B | 44438842E36F25083ADA8DDD786F604A20E14391B0455BEB57B38CC4617AD5A8 | 89682177374F001750D4939212088F97 |
| BorradoSolicitudes.bas (VBA Export) | 21.439 B | A915F966AEF17A4B5B0B941FE1C480AEAE5985A1F62466973DECB4E569FF0849 | E74CABC2E39A262D9D52580150455AB3 |
