# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-11-06 | Hito Oficial: Versión de Referencia 2026-11-06 00:45
### Ámbito: Fortalecimiento Quirúrgico Integral de Bloqueos y Ocultación de Habitaciones en Access (OP-FORT-01..05 y OP-FINAL-BLOQ-01..03 con Control Optimista de Concurrencia, Transacciones ACID Unificadas, Serialización Inyectiva Canónica y Blindaje Anti-Lost Update)

---

### 1. Resumen de la Intervención y Contexto Técnico

En las sesiones de trabajo del **05 y 06 de Noviembre de 2026**, se completó con éxito la auditoría forense, implementación y consolidación del subsistema de **Persistencia y Bloqueo de Habitaciones y Ocultación en Calendarios en Microsoft Access (`Residencia_BE.accdb`)**, estructurado en dos grandes fases quirúrgicas:
1. **Fase de Fortalecimiento Base (`OP-FORT-01` a `OP-FORT-05`):** Creación del esquema relacional en Access (tablas `BloqueosHabitaciones` y `EstadoHabitacionesCalendario`), migración de datos locales a BD, lectura al inicio y guardado al salir.
2. **Fase de Blindaje Arquitectónico Final (`OP-FINAL-BLOQ-01` a `OP-FINAL-BLOQ-03`):** Subsanación exhaustiva de los 7 hallazgos arquitectónicos de concurrencia y transaccionalidad detectados en la auditoría técnica independiente (H-01 a H-07).

---

### 2. Trazabilidad Estricta: Funcionalidades vs. Copias de Seguridad Incrementales

| Copia Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Hito / Funcionalidad Técnica Vinculada |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.xlsm` | 6.616.947 B | `06E31C80A08574CA44F016850FBA6655` | `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9` | **Línea Base Previa de Partida.** Incorpora Resolución Dinámica `DB_PATH`, Purga WAL SICO-N2, Protección `UserInterfaceOnly` y Coherencia Habitacional. |
| `Residencia_BE_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.accdb` | 6.754.304 B | `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | Respaldo de partida de la base de datos Access previa a la creación de las tablas de bloqueos. |
| `BDAS_Multiusuario_BACKUP_2026-11-05_2251_PRE_FORT_BLOQ.xlsm` | 6.654.465 B | `04D537D0611C693689C6085ACEA06724` | `A5E5884783B75149238A9DB6C1793A8CE3B50AA505E346EDAEC7925B2BD4C18A` | **Punto de Control Pre-Fortalecimiento.** Respaldo obligatorio previo a la aplicación de las operaciones base OP-FORT-01 a OP-FORT-05. |
| `Residencia_BE_BACKUP_2026-11-05_2251_PRE_FORT_BLOQ.accdb` | 6.754.304 B | `FDCEF694DDB8F06C610768483C6E14A9` | `4E8DEE59F74045257B3B25CE1BB9A1F512F0D55FDC00841BA6A0A41259E8407D` | Respaldo previo de Access antes de crear las tablas `BloqueosHabitaciones` y `EstadoHabitacionesCalendario`. |
| `BDAS_Multiusuario_BACKUP_2026-11-05_PRE_FINAL_BLOQ.xlsm` | 6.557.819 B | `50C5F1B739E997CEC92ED96E5762F5D7` | `69A604E6E2DD83526E0D5126CCC06C8C0701171B00D2AFFB23AD1D3D68D6F249` | **Punto de Control Pre-Final.** Respaldo con OP-FORT-01..05 ya aplicadas e integradas, previo a la subsanación de los 7 hallazgos H-01 a H-07. |
| `BDAS_Multiusuario_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.xlsm` | 6,590,683 B | `CD9A878DB02B874BF17A4F2F9EDA3BA6` | `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F` | **Versión de Referencia Oficial Consolidada.** Cierre completo con OP-FINAL-BLOQ-01, 02 y 03 integradas, 173 componentes VBE, 0 errores y cero mojibake. |
| `Residencia_BE_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.accdb` | 6,754,304 B | `5816C92401274DD2A84EB238AB92980E` | `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40` | **Base de Datos Referencia Oficial Consolidada.** Incorpora tablas `BloqueosHabitaciones` y `EstadoHabitacionesCalendario` con índices únicos y esquemas blindados. |

---

### 3. Detalle Técnico de Funcionalidades y Mejoras Implementadas

#### A. Hallazgo H-01: Control Optimista de Concurrencia y Detección de *Lost Updates*
- **Componentes modificados:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L5800) y [`ThisWorkbook.cls`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls#L340).
- **Problema previo:** Si el Puesto A abría la hoja de bloqueos y mientras tanto el Puesto B guardaba un bloqueo nuevo en Access, cuando el Puesto A salía de la hoja ejecutaba `DELETE FROM BloqueosHabitaciones` + `INSERT`, destruyendo el bloqueo guardado por B.
- **Solución implementada:**
  - Nueva función `CalcularFirmaBloqueosDesdeBD(cn, claveRes)`.
  - En `GuardarBloqueosEnBD`, antes de escribir, se compara la firma actual en Access con la firma base que el operador cargó al entrar (`s_DictFirmas(claveRes)`).
  - Si difieren, la transacción se aborta con `RollbackTrans` devolviendo `False` y mensaje de colisión concurrente.
  - En `ThisWorkbook.cls` (`Workbook_SheetDeactivate`), si salta colisión concurrente, se ofrece al usuario recargar de inmediato la información de Access sin destruir los datos del otro operador.

#### B. Hallazgo H-02: Transacciones ACID Unificadas en Guardado de Sede
- **Componente modificado:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L5980).
- **Problema previo:** `GuardarBloqueosYEstadoHabitacionesEnBD` abría hasta 16 transacciones aisladas (1 para bloqueos y 15 para filas del calendario). Una caída de red a mitad dejaba la base de datos en estado híbrido inconsistente.
- **Solución implementada:**
  - Se engloba todo el proceso de la sede en una **única transacción atómica** (`cn.BeginTrans` ... `cn.CommitTrans`).
  - Si falla cualquier fila del calendario o la persistencia de bloqueos, `cn.RollbackTrans` revierte toda la operación en bloque, preservando la consistencia relacional.

#### C. Hallazgo H-03: Notificación Explícita de Errores de Recarga tras Sincronización
- **Componente modificado:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3321) y ([L3580](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3580)).
- **Problema previo:** En `SincronizarHojaDesdeAccess`, las dos llamadas a `CargarBloqueosYEstadoHabitacionesDesdeBD` descartaban el valor de retorno booleano, ocultando fallos de lectura de bloqueos al operador.
- **Solución implementada:**
  - Se captura el retorno `Boolean` y la variable `errBloqSync`. Si devuelve `False`, se muestra un mensaje informativo `MsgBox` con el detalle del fallo sin abortar el reprotegido de la hoja.

#### D. Hallazgo H-04: Importación Atómica Multisede con Instantánea de Reversión
- **Componente modificado:** [`CopiarDatosHojasGijonSotoOviedo.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/CopiarDatosHojasGijonSotoOviedo.bas#L930).
- **Problema previo:** La condición `OR` podía considerar exitosa la importación si faltaba una hoja, y no había mecanismo de restauración si saltaba error a mitad del copiado.
- **Solución implementada:**
  - Se exige la existencia de las 4 hojas involucradas con condición `AND` estricta (`wsBloqOrig`, `wsBloqDest`, `wsCalOrig`, `wsCalDest`).
  - Se toma una instantánea previa en memoria (`arrBackupBloqDest`, `estadosBackupCalDest`). Ante cualquier error de ejecución, se restauran íntegramente las hojas de destino a su estado previo antes de salir.

#### E. Hallazgo H-05: Verificación Profunda de Esquemas e Índices Reales
- **Componente modificado:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L5560).
- **Solución implementada:**
  - Comprobación de que el índice único `idx_EstHab_ResFila` no contiene una tercera columna espuria (`colExtra`).
  - Verificación de la existencia del índice `idx_Bloq_Res` y validación de tipos DAO/ADO (`DefinedSize >= 20/50/255`).

#### F. Hallazgo H-06: Validación Preventiva sin Truncado Silencioso
- **Componente modificado:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L5710).
- **Solución implementada:**
  - Se añade en `ValidarHojaBloqueosAntesDeGuardar` la comprobación explícita `Len(sHab) > 50` y `Len(sMot) > 255`. Si se supera, se rechaza la operación informando la fila y longitud para que el usuario corrija sin pérdida de caracteres.

#### G. Hallazgo H-07: Serialización Canónica Inyectiva con Prefijos de Longitud
- **Componente modificado:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L5460).
- **Solución implementada:**
  - Sustitución de DJB2 y SHA-256 COM por serialización canónica exacta con prefijos de longitud por celda (`Len(v) & ":" & v & "|"`): **0% riesgo de colisión matemática**, inmune a `Error 6 (Desbordamiento)` en Excel 32 bits y 0 dependencias COM externas.

---

### 4. Certificación de Calidad y Puertas de Aceptación

- [x] **Compilación VBE:** `Application.VBE.CommandBars.FindControl(1, 578).Enabled == False` (0 errores).
- [x] **Total Componentes VBE:** Exactamente 173 componentes (170 componentes no objetivo 100% inalterados).
- [x] **Codificación Nativa:** Windows-1252 ANSI estricta en todos los módulos exportados a `vba-modules/`.
- [x] **Anti-Mojibake:** 0 patrones anómalos detectados en el código VBA.
- [x] **Regla Estricta de Nombres:** Preservación de 'GIJÓN' como literal estricto.
- [x] **Inmutabilidad de Golden Master:** La carpeta `H:\ResidenciaBD\Plantilla\GoldenMaster\` no ha sido modificada ni accedida.
