# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-11-02 | Hito Oficial Consolidado: Versión de Referencia 2026-11-02 03:30
### Ámbito: Cierre SICO-N2 WAL Sin Bloqueo, Optimización Anti-Bloqueo DWM en GIJÓN y SOTO, Gestión Integral de Borrado RGPD (Manual y Automático en Arranque), Reintegración Quirúrgica Oviedo y Blindaje Anti-Error 3704 en Gestión de Usuarios

---

## 1. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia exacta entre cada hito técnico de implementación en la jornada 2026-11-02, los archivos de respaldo incrementales generados en disco, sus dimensiones, hashes criptográficos y el conjunto de funcionalidades y correcciones vinculadas:

| Hito / Copia Incremental | Ruta del Archivo de Respaldo | Tamaño (Bytes) | Hashes Criptográficos (MD5 / SHA-256) | Funcionalidades, Operaciones y Módulos Vinculados |
| :--- | :--- | :--- | :--- | :--- |
| **0. Estado Base de Partida**<br>*(Hito 2026-11-01 04:25)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-11-01_0425.accdb` | 6.809.197 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `5727B66025C7278C95B5596CA22BD351`<br>**XLSM SHA256:** `E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Punto de partida de la sesión. Incorporaba Perfeccionamiento de Concurrencia V5.2, monotonicidad `idTransBase`, verificación de Lease Lock (120s), discriminación de fórmulas pre-arrastradas y blindaje de eventos. |
| **1. Backup Preventivo PRE-CIERRE SICO-N2**<br>*(Antes de Cierre WAL y DWM)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0025_PRE_CIERRE_SICO_N2.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_0025_PRE_CIERRE_SICO_N2.accdb` | 6.809.315 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `F6C49DB88A5A885D429B9984D01B56CA`<br>**XLSM SHA256:** `21747171858A9E2F79BE3E53DA6E5C5A3751EE4DD3720A92A3F71F2EF65EB30A`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Copia de seguridad preventiva generada antes de abordar el cierre definitivo del encolador offline SICO-N2 y la optimización de borrado total contra congelamiento DWM en GIJÓN y SOTO. |
| **2. Backup POST-CIERRE SICO-N2**<br>*(Cierre WAL y Optimización DWM)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0025_POST_CIERRE_SICO_N2.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-11-02_0025_POST_CIERRE_SICO_N2.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_0025_POST_CIERRE_SICO_N2.accdb` | 6.825.806 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `B460BCFE65EF7B9A6F96993C4D965791`<br>**XLSM SHA256:** `9E0BB5192C4582B592418B6F89F025C4F5BA9E81FBA782897DBC31B0C370A227`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Incorpora: 1) Desbloqueo del WAL en arranque (`ModuloSeguridadLocal.bas`), eliminando `If ultFila >= 2 Then Exit Sub` y procesando `ELIMINAR_ORDEN` matricialmente en RAM. 2) Supresión de `.Delete` sobre 16.348 columnas en `BorradoDatosAlojamiento.bas` (usando `.Clear` acotado al UsedRange real) y recálculo escalonado con `EnableCalculation = False` en Calendarios y Resumen junto con `DoEvents`. |
| **3. Backup Preventivo PRE-BORRADO RGPD**<br>*(Antes de Gestión Borrado)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0110_PRE_GESTION_BORRADO_RGPD.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_0110_PRE_GESTION_BORRADO_RGPD.accdb` | 6.825.806 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `B460BCFE65EF7B9A6F96993C4D965791`<br>**XLSM SHA256:** `9E0BB5192C4582B592418B6F89F025C4F5BA9E81FBA782897DBC31B0C370A227`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Respaldo previo a la implementación de la evaluación en memoria de `FechaSalida` (Col M) / `FechaPeticion` (Col B), compactación física atómica en lotes de hasta 100 filas y automatización de arranque. |
| **4. Backup Consolidación POST-BORRADO RGPD**<br>*(Cierre Auditoría RGPD)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0110_POST_GESTION_BORRADO_RGPD.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-11-02_0110_POST_GESTION_BORRADO_RGPD.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_0110_POST_GESTION_BORRADO_RGPD.accdb` | 6.838.431 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `33AC497F6CF682AFCB79DDAAA07446C2`<br>**XLSM SHA256:** `8EE9DB2D367F0E7C4FC1E82BB23423C925264B7291F67D020A4100B93BCFA0C6`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Incorpora: 1) `modDatabase.bas`: `ObtenerFilasYOrdenesArchivablesSede`, `ContarOrdenesArchivablesBD` y `ArchivarOrdenesAntiguasBD` evaluando Col M/B y compactando con `Application.Union` + `.Delete Shift:=xlUp` en lotes de $\le 100$. 2) `ModuloArchivadoDatosBDAS.bas`: Supresión de llamadas redundantes a `SincronizarHojaDesdeAccess`. 3) `ThisWorkbook.cls`: Conexión de `ArchivarDatosAntiguosAlAbrir` en Paso 4 de `Workbook_Open`. *(Promovido a Ref. 01:20)*. |
| **5. Backup Preventivo PRE-REINTEGRAR OVIEDO**<br>*(Tras Pruebas Manuales de Borrado)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0228_PRE_REINTEGRAR_OVIEDO.xlsm` | 6.332.879 B (XLSM) | **XLSM MD5:** `B2461A0ABEA664BDD56BAE5AD3EA1248`<br>**XLSM SHA256:** `9582BEBF64FA140AA4F4C4E0C741E98CC63A38860E190DDC4D456733289755DC` | Respaldo del libro activo tras la prueba de borrado manual con 29/30 días realizada por el usuario en `RESIDENCIA OVIEDO`, previo a la reintegración quirúrgica. |
| **6. Backup POST-REINTEGRAR OVIEDO Y SANEAMIENTO WAL** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0228_POST_REINTEGRAR_OVIEDO.xlsm` | 6.843.327 B (XLSM) | **XLSM MD5:** `BB9A7FA812D939647C76AB52A317C53E`<br>**XLSM SHA256:** `F52FA5DF4968D51BDB7ADED7862BD97F5F2BCB5A14059EF617D3FF8414922C52` | Reintegración quirúrgica de las 1.802 órdenes en `RESIDENCIA OVIEDO` (rango A1:AH4000), recomposición de `RESUMEN OVIEDO`, restablecimiento de `SYS_COLA_SYNC` y purga de 1.795 operaciones `ELIMINAR_ORDEN` en el archivo físico `.wal`. |
| **7. Versión de Referencia Consolidada Final**<br>*(Hito Oficial 2026-11-02 03:30)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm`<br>`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.accdb` | 6.430.806 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `F73F5C628319611FD8C56CA773E60B67`<br>**XLSM SHA256:** `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | **Versión de Referencia Vigente Oficial.** Incorpora todas las optimizaciones anteriores más el blindaje de Gestión de Usuarios en `modDatabase.bas` (sonda preventiva SICO-N2 en `ObtenerListaUsuariosDelegada` y eliminación de Error 3704 sobre objetos cerrados en `ErrorHandler` de `ObtenerListaUsuariosDelegada` y `ConsultarUsuarioAccess`). 0 errores de compilación COM. |

---

## 2. Descripción Técnica de las Modificaciones Implementadas

### 2.1. Blindaje Anti-Error 3704 y Sonda SICO-N2 en Gestión de Usuarios (`modDatabase.bas`)
- **Diagnóstico del Fallo:** Al pulsar `HERRAMIENTAS -> Gestión de usuarios` y seleccionar la opción 5 (*"Ver todos los usuarios del sistema"*), la aplicación saltaba al depurador con el error en tiempo de ejecución 3704 (*"La operación no está permitida si el objeto está cerrado"*), resaltando `: cn.Close` en la rutina `ErrorHandler:` de `ObtenerListaUsuariosDelegada`.
- **Causas Técnicas Identificadas:**
  1. *Ausencia de comprobación de estado ADO:* Si `cn.Open` fallaba (por estar sin red o sin enlace con la base de datos Access), `cn` estaba instanciada pero en estado `adStateClosed` (`0`). Al entrar al manejador de errores, `If Not cn Is Nothing Then cn.Close` intentaba cerrar una conexión no abierta, lo cual genera de inmediato el error 3704.
  2. *Inoperancia de `On Error Resume Next` en manejador activo:* En VBA, cuando el flujo se encuentra dentro de un bloque `ErrorHandler:`, el entorno está en modo de excepción activa; cualquier `On Error Resume Next` en línea es ignorado por el motor de ejecución, convirtiendo el error secundario 3704 en un fallo no capturado que detenía la macro en el depurador.
  3. *Falta de sonda de red preventiva SICO-N2:* La función intentaba ejecutar `cn.Open` a ciegas, sin comprobar previamente si el equipo se encontraba en contingencia offline.
- **Solución Implementada:**
  - En `ObtenerListaUsuariosDelegada`:
    - Incorporación de la sonda de conectividad previa:
      ```vba
      If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then
          ObtenerListaUsuariosDelegada = "No se puede consultar la lista de usuarios en este momento." & vbCrLf & vbCrLf & _
                                         "La base de datos Access no está accesible (Modo Contingencia SICO-N2 activo)."
          Exit Function
      End If
      ```
    - Sustitución del cierre ciego por comprobación explícita de `adStateOpen`:
      ```vba
      ErrorHandler:
          On Error Resume Next
          If Not rs Is Nothing Then
              If rs.State = adStateOpen Then rs.Close
              Set rs = Nothing
          End If
          If Not cn Is Nothing Then
              If cn.State = adStateOpen Then cn.Close
              Set cn = Nothing
          End If
          On Error GoTo 0
          ObtenerListaUsuariosDelegada = "Error al consultar usuarios en Access: " & Err.Description
      End Function
      ```
  - En `ConsultarUsuarioAccess`: Aplicación idéntica del blindaje de cierre seguro en `ErrorHandler:` verificando `adStateOpen`.

### 2.2. Reintegración Quirúrgica de Registros en `RESIDENCIA OVIEDO` y Saneamiento WAL SICO-N2
- **Diagnóstico:** Tras la prueba de archivado manual con 29 días, la hoja `RESIDENCIA OVIEDO` redujo sus filas a 2.646 (448 órdenes). La investigación reveló que la función `DIAS_RETENCION()` forzaba un valor por defecto de 90 días si el parámetro era menor a 30 (`If dias >= 30 And dias <= 365`), archivando órdenes con salida $\le 04/10/2026$ y encolando 1.795 operaciones `ELIMINAR_ORDEN` en la cola WAL de contingencia N2.
- **Solución Implementada:**
  - Restauración quirúrgica completa de la cuadrícula `A1:AH4000` en `RESIDENCIA OVIEDO` desde la versión de referencia oficial, recuperando íntegramente las 1.802 órdenes (filas 2 a 4000), incluyendo las órdenes 108 y 109, con fórmulas, validaciones y formato nativo.
  - Recomposición de `RESUMEN OVIEDO` (`A1:R1803`) alineando los totales con el dataset de 1.802 expedientes.
  - Saneamiento preventivo de `SYS_COLA_SYNC` y del archivo físico local `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\SYS_COLA_SYNC_SALÓN-PC.wal` (depurado de 2.133 a 336 líneas), eliminando las operaciones residuales `ELIMINAR_ORDEN` de Oviedo para evitar re-ejecuciones involuntarias al abrir Excel o sincronizar.

### 2.3. Cierre Definitivo SICO-N2 WAL (`ModuloSeguridadLocal.bas`)
- **Solución Implementada:**
  - Supresión definitiva de la cláusula de salida anticipada `If ultFila >= 2 Then Exit Sub`, asegurando que el archivo WAL en disco sea la fuente no volátil de verdad.
  - Carga masiva de la cola a memoria RAM en una sola lectura COM (`arrCola = wsCola.Range(...).Value`).
  - En la rutina `RehidratarYReaplicarEnHojasOffline`, procesamiento en bloque de todas las operaciones pendientes `ELIMINAR_ORDEN` (parciales con `Application.Union` + `.Delete Shift:=xlUp` y totales vaciando datos preservando fórmulas y formato condicional).

### 2.4. Optimización Anti-Bloqueo DWM (`BorradoDatosAlojamiento.bas`)
- **Solución Implementada:**
  - Supresión del comando destructivo `.Delete` sobre 16.348 columnas vacías de Excel; sustituido por `.Clear` acotado exclusivamente si el `UsedRange` real excede la columna 36 (`If ultColUsada > 36 Then ws.Range(...).Clear`).
  - Recálculo escalonado mediante desactivación temporal de `EnableCalculation` en `wsCal` y `wsSum` junto con llamadas estratégicas a `DoEvents` para ceder ciclos de CPU al sistema operativo antes de reactivar el cálculo, erradicando el congelamiento de interfaz ("Excel no responde") en GIJÓN y SOTO.

### 2.5. Evaluación Matricial y Compactación Atómica RGPD (`modDatabase.bas`)
- **Solución Implementada:**
  - Implementación de `ObtenerFilasYOrdenesArchivablesSede`: evaluación matricial en RAM priorizando Columna M (`FechaSalida`) con fallback a Columna B (`FechaPeticion`).
  - `ContarOrdenesArchivablesBD`: cálculo integrado en Access y soporte en contingencia offline.
  - `ArchivarOrdenesAntiguasBD`: delegación transaccional en `EliminarOrdenesLoteBD` (Access y WAL) y compactación física inmediata de las hojas de residencia mediante `Application.Union` + `.Delete Shift:=xlUp` en sub-lotes de $\le 100$ filas.

### 2.6. ModuloArchivadoDatosBDAS.bas y ThisWorkbook.cls
- **Solución Implementada:**
  - Eliminación de la invocación redundante a `SincronizarHojaDesdeAccess` en `ArchivarDatosManual` y `ArchivarDatosAntiguosAlAbrir`.
  - Inyección de `Call ModuloArchivadoDatosBDAS.ArchivarDatosAntiguosAlAbrir` en el Paso 4 del evento `Workbook_Open` en `ThisWorkbook.cls`.

---

## 3. Estado de Compilación y Validación de Integridad

- **Compilación VBA COM:** 100% limpia sin errores (`VBE.CommandBars.FindControl(1, 578).Enabled == False`).
- **Codificación:** Windows-1252 ANSI sin BOM en todos los componentes exportados.
- **Toponimia:** Cumplimiento riguroso del literal `GIJÓN` / `GIJON`.
- **Golden Master:** Réplica protegida inmutable e intacta en `H:\ResidenciaBD\Plantilla\GoldenMaster\`.
