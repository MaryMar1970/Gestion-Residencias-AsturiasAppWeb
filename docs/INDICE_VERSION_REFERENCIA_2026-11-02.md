# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: 2026-11-02 03:30 (Cierre SICO-N2 WAL + Optimización Anti-Bloqueo DWM en GIJÓN y SOTO + Gestión Borrado RGPD + Reintegración Oviedo + Blindaje Error 3704 en Gestión de Usuarios)

---

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  *Tamaño:* 6,430,806 bytes | *MD5:* `F73F5C628319611FD8C56CA773E60B67` | *SHA-256:* `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`

- **Copia Timestamped en VersionReferencia (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm`
  *Tamaño:* 6,430,806 bytes | *MD5:* `F73F5C628319611FD8C56CA773E60B67` | *SHA-256:* `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`

- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
  *Tamaño:* 6,746,112 bytes | *MD5:* `6237949FA8E9F503C9623F0A7294310D` | *SHA-256:* `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196`

- **Copia Timestamped Base de Datos (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-02_0330.accdb`
  *Tamaño:* 6,746,112 bytes | *MD5:* `6237949FA8E9F503C9623F0A7294310D` | *SHA-256:* `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196`

- **Copia Activa de Trabajo (Desarrollo y Pruebas):**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`
  *Tamaño:* 6,430,806 bytes | *MD5:* `F73F5C628319611FD8C56CA773E60B67` | *SHA-256:* `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`

- **Copia Desplegada en Red (Producción Multiusuario Operativa):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`
  *Tamaño:* 6,430,806 bytes | *MD5:* `F73F5C628319611FD8C56CA773E60B67` | *SHA-256:* `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`

---

### 2. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia estricta entre las intervenciones técnicas realizadas en la jornada 2026-11-02 y las copias de seguridad incrementales generadas en `H:\ResidenciaBD\BackupsDiarios\`:

| Backup Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Finalidad Técnica y Alcance Operativo |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0025_PRE_CIERRE_SICO_N2.xlsm` | 6.809.315 B | `F6C49DB88A5A885D429B9984D01B56CA` | `21747171858A9E2F79BE3E53DA6E5C5A3751EE4DD3720A92A3F71F2EF65EB30A` | **Línea Base Previa.** Estado anterior al cierre SICO-N2 y optimización DWM. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0025_POST_CIERRE_SICO_N2.xlsm` | 6.825.806 B | `B460BCFE65EF7B9A6F96993C4D965791` | `9E0BB5192C4582B592418B6F89F025C4F5BA9E81FBA782897DBC31B0C370A227` | **Cierre WAL y DWM.** Desbloqueo de cola WAL (`ModuloSeguridadLocal.bas`) y supresión de `.Delete` sobre 16k columnas con recálculo escalonado (`BorradoDatosAlojamiento.bas`). |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0110_PRE_GESTION_BORRADO_RGPD.xlsm` | 6.825.806 B | `B460BCFE65EF7B9A6F96993C4D965791` | `9E0BB5192C4582B592418B6F89F025C4F5BA9E81FBA782897DBC31B0C370A227` | **Punto de Control Pre-RGPD.** Respaldo previo a la implementación de la evaluación por Col M/B y compactación en lotes. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0110_POST_GESTION_BORRADO_RGPD.xlsm` | 6.838.431 B | `33AC497F6CF682AFCB79DDAAA07446C2` | `8EE9DB2D367F0E7C4FC1E82BB23423C925264B7291F67D020A4100B93BCFA0C6` | **Gestión Borrado RGPD Completa.** Incorpora evaluación Col M/B, delegación transaccional en `EliminarOrdenesLoteBD`, compactación física por lotes ($\le 100$ filas) en `modDatabase.bas`, saneamiento de `ModuloArchivadoDatosBDAS.bas` y conexión en `Workbook_Open` de `ThisWorkbook.cls`. *(Promovido previamente a Referencia 01:20)*. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0228_PRE_REINTEGRAR_OVIEDO.xlsm` | 6.332.879 B | `B2461A0ABEA664BDD56BAE5AD3EA1248` | `9582BEBF64FA140AA4F4C4E0C741E98CC63A38860E190DDC4D456733289755DC` | **Pre-Reintegración Oviedo.** Respaldo preventivo antes de reintegrar quirúrgicamente los 1.802 registros de Oviedo y depurar WAL. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0228_POST_REINTEGRAR_OVIEDO.xlsm` | 6.843.327 B | `BB9A7FA812D939647C76AB52A317C53E` | `F52FA5DF4968D51BDB7ADED7862BD97F5F2BCB5A14059EF617D3FF8414922C52` | **Post-Reintegración Oviedo y Saneamiento WAL.** Restablecimiento de `RESIDENCIA OVIEDO` (1.802 órdenes / 4.000 filas), recomposición de `RESUMEN OVIEDO` y purga de 1.795 operaciones `ELIMINAR_ORDEN` en cola WAL. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.xlsm` | 6.430.806 B | `F73F5C628319611FD8C56CA773E60B67` | `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60` | **Versión de Referencia Consolidada Final.** Incorpora el blindaje de Gestión de Usuarios en `modDatabase.bas` (sonda SICO-N2 preventiva en `ObtenerListaUsuariosDelegada` y eliminación de Error 3704 sobre objetos cerrados en `ErrorHandler`). |
| `Residencia_BE_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.accdb` | 6.746.112 B | `6237949FA8E9F503C9623F0A7294310D` | `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | **Respaldo Central de Base de Datos Access.** Estructura de tablas `Usuarios`, `Ordenes` y `Ordenes_Historico` 100% íntegra. |

- **Histórico de Referencias Previas en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm` (6.430.806 B | MD5: `F73F5C628319611FD8C56CA773E60B67`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0110.xlsm` (6.838.431 B | MD5: `33AC497F6CF682AFCB79DDAAA07446C2`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm` (6.809.197 B | MD5: `5727B66025C7278C95B5596CA22BD351`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0345.xlsm` (6.792.119 B | MD5: `D616F3366D4AB18F507B5AE47A49DEE7`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_2015.xlsm` (6.749.596 B | MD5: `6AE80126541D7D7F273B94929ECAE6A9`)

---

### 3. Inventario de Componentes y Mejoras Clave del Hito 2026-11-02 03:30

1. **`modDatabase.bas` (Blindaje Anti-Error 3704 y Sonda SICO-N2 en Gestión de Usuarios):**
   - En `ObtenerListaUsuariosDelegada`:
     * Comprobación preventiva de conectividad con `If Not ModuloSeguridadLocal.SondaRedDisponible(DB_PATH) Then`: al operar offline o sin enlace a la base de datos Access, devuelve un aviso limpio e informativo sin forzar la apertura de conexiones.
     * Reestructuración del bloque `ErrorHandler:` para verificar `If rs.State = adStateOpen Then rs.Close` e `If cn.State = adStateOpen Then cn.Close` antes de liberar los punteros, erradicando el error en tiempo de ejecución 3704 ("La operación no está permitida si el objeto está cerrado") al depurar.
   - En `ConsultarUsuarioAccess`:
     * Reestructuración idéntica de `ErrorHandler:` garantizando que jamás se invoque `.Close` sobre conexiones o recordsets cerrados.

2. **`modDatabase.bas` (Evaluación Matricial y Compactación Atómica RGPD):**
   - Implementación de `ObtenerFilasYOrdenesArchivablesSede`: evaluación matricial en RAM priorizando Columna M (`FechaSalida`) con fallback a Columna B (`FechaPeticion`).
   - `ContarOrdenesArchivablesBD`: cálculo integrado en Access y soporte en contingencia offline.
   - `ArchivarOrdenesAntiguasBD`: delegación transaccional en `EliminarOrdenesLoteBD` (Access y WAL) y compactación física inmediata de las hojas de residencia mediante `Application.Union` + `.Delete Shift:=xlUp` en sub-lotes de $\le 100$ filas.

3. **`ModuloSeguridadLocal.bas` (Cierre SICO-N2 y Desbloqueo WAL):**
   - Eliminación de la guarda prematura `If ultFila >= 2 Then Exit Sub` en `RehidratarColaDesdeDiscoWALSiNecesario`.
   - Lecturas y escrituras matriciales en RAM (`arrCola = wsCola.Range(...).Value`) para operaciones WAL sin sobrecarga COM celda por celda.
   - Enrutamiento masivo de operaciones `ELIMINAR_ORDEN` en `RehidratarYReaplicarEnHojasOffline` (parcial y total).
   - Depuración y saneamiento preventivo del archivo físico WAL local, eliminando operaciones residuales generadas durante pruebas.

4. **`BorradoDatosAlojamiento.bas` (Optimización Anti-Bloqueo DWM):**
   - Supresión del comando destructivo `.Delete` sobre 16.348 columnas vacías de Excel; sustituido por `.Clear` acotado exclusivamente si `ultColUsada > 36`.
   - Recálculo escalonado mediante suspensión temporal de `EnableCalculation` en `wsCal` y `wsSum` junto con llamadas estratégicas a `DoEvents`, erradicando el congelamiento de interfaz ("Excel no responde") en GIJÓN y SOTO.

5. **`ModuloArchivadoDatosBDAS.bas` (Saneamiento de Flujos):**
   - Eliminación de la invocación redundante a `SincronizarHojaDesdeAccess` en `ArchivarDatosManual` y `ArchivarDatosAntiguosAlAbrir`.

6. **`ThisWorkbook.cls` (Disparo Automático en Arranque):**
   - Conexión de `Call ModuloArchivadoDatosBDAS.ArchivarDatosAntiguosAlAbrir` en el Paso 4 del evento `Workbook_Open`, automatizando el cumplimiento RGPD sin intervención del usuario.

---

### 4. Garantías de Calidad y Gobernanza Técnica

- **Compilación VBA COM:** 100% verificada en instancia real de Excel COM; `VBE.CommandBars.FindControl(1, 578).Enabled == False` (**0 errores de compilación**).
- **Codificación Nativa:** Windows-1252 ANSI sin BOM en todos los módulos exportados en `vba-modules\`.
- **Control Anti-Mojibake:** 0 caracteres corruptos detectados en todo el repositorio.
- **Regla Estricta de Toponimia:** Uso exclusivo del literal `GIJÓN` / `GIJON`. Prohibición total de escapes numéricos (`Chr(211)`).
- **Inmutabilidad del Golden Master:** La réplica `H:\ResidenciaBD\Plantilla\GoldenMaster\` no ha sido modificada.
