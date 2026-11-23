# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: Hito 2026-11-06 00:45 (Fortalecimiento Quirúrgico Integral de Bloqueos y Ocultación de Habitaciones en Access (OP-FORT-01..05 y OP-FINAL-BLOQ-01..03 con Control Optimista de Concurrencia, Transacciones ACID Unificadas, Serialización Inyectiva Canónica y Blindaje Anti-Lost Update))

---

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`  
  *Tamaño:* 6,590,683 bytes | *MD5:* `CD9A878DB02B874BF17A4F2F9EDA3BA6` | *SHA-256:* `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F`

- **Copia Timestamped en VersionReferencia (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-06_0045.xlsm`  
  *Tamaño:* 6,590,683 bytes | *MD5:* `CD9A878DB02B874BF17A4F2F9EDA3BA6` | *SHA-256:* `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F`

- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`  
  *Tamaño:* 6,754,304 bytes | *MD5:* `5816C92401274DD2A84EB238AB92980E` | *SHA-256:* `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40`

- **Copia Timestamped Base de Datos (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-06_0045.accdb`  
  *Tamaño:* 6,754,304 bytes | *MD5:* `5816C92401274DD2A84EB238AB92980E` | *SHA-256:* `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40`

- **Copia Activa de Trabajo (Desarrollo y Pruebas):**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6,590,683 bytes | *MD5:* `CD9A878DB02B874BF17A4F2F9EDA3BA6` | *SHA-256:* `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F`

- **Copia Desplegada en Red (Producción Multiusuario Operativa):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6,590,683 bytes | *MD5:* `CD9A878DB02B874BF17A4F2F9EDA3BA6` | *SHA-256:* `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F`

---

### 2. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia estricta entre las intervenciones técnicas realizadas en las jornadas 2026-11-05 / 2026-11-06 y las copias de seguridad incrementales generadas en `H:\ResidenciaBD\BackupsDiarios\`:

| Backup Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Finalidad Técnica y Alcance Operativo |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.xlsm` | 6.616.947 B | `06E31C80A08574CA44F016850FBA6655` | `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9` | **Línea Base Previa de Partida.** Estado anterior a la implementación de persistencia de bloqueos. |
| `Residencia_BE_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.accdb` | 6.754.304 B | `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | Respaldo previo de base de datos Access de partida. |
| `BDAS_Multiusuario_BACKUP_2026-11-05_2251_PRE_FORT_BLOQ.xlsm` | 6.654.465 B | `04D537D0611C693689C6085ACEA06724` | `A5E5884783B75149238A9DB6C1793A8CE3B50AA505E346EDAEC7925B2BD4C18A` | **Punto de Control Pre-Fortalecimiento.** Respaldo previo a aplicar OP-FORT-01 a OP-FORT-05. |
| `Residencia_BE_BACKUP_2026-11-05_2251_PRE_FORT_BLOQ.accdb` | 6.754.304 B | `FDCEF694DDB8F06C610768483C6E14A9` | `4E8DEE59F74045257B3B25CE1BB9A1F512F0D55FDC00841BA6A0A41259E8407D` | Respaldo previo de base de datos Access antes de crear el esquema de bloqueos. |
| `BDAS_Multiusuario_BACKUP_2026-11-05_PRE_FINAL_BLOQ.xlsm` | 6.557.819 B | `50C5F1B739E997CEC92ED96E5762F5D7` | `69A604E6E2DD83526E0D5126CCC06C8C0701171B00D2AFFB23AD1D3D68D6F249` | **Punto de Control Pre-Final.** Respaldo con OP-FORT-01..05 ya aplicadas, previo a resolver H-01 a H-07. |
| `BDAS_Multiusuario_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.xlsm` | 6,590,683 B | `CD9A878DB02B874BF17A4F2F9EDA3BA6` | `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F` | **Versión de Referencia Oficial Consolidada.** Copia de cierre con OP-FINAL-BLOQ-01..03 integradas. |
| `Residencia_BE_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.accdb` | 6,754,304 B | `5816C92401274DD2A84EB238AB92980E` | `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40` | **Base de Datos Referencia Oficial Consolidada.** Tablas `BloqueosHabitaciones` y `EstadoHabitacionesCalendario` activas y verificadas. |

- **Histórico de Referencias Previas en `H:\\ResidenciaBD\\Plantilla\\VersionReferencia\\Historico\\`:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-03_0015.xlsm` (6.616.947 B | MD5: `06E31C80A08574CA44F016850FBA6655`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_2202.xlsm` (6.385.789 B | MD5: `549921707952BC6EAA423A6CBC107500`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm` (6.430.806 B | MD5: `F73F5C628319611FD8C56CA773E60B67`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0110.xlsm` (6.838.431 B | MD5: `33AC497F6CF682AFCB79DDAAA07446C2`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm` (6.809.197 B | MD5: `5727B66025C7278C95B5596CA22BD351`)

---

### 3. Inventario de Componentes y Mejoras Clave del Hito 2026-11-06 00:45

1. **`modDatabase.bas` (Control Optimista de Concurrencia en Bloqueos):**
   - Comparación canónica dentro de la transacción ADO mediante `CalcularFirmaBloqueosDesdeBD(cn, claveRes)`.
   - Detección de modificaciones concurrentes realizadas por otros puestos mientras la hoja estaba abierta, evitando el *Lost Update*.

2. **`modDatabase.bas` (Transacciones ACID Unificadas):**
   - Englobamiento de la persistencia de `BloqueosHabitaciones` y todas las filas de `EstadoHabitacionesCalendario` en un solo `cn.BeginTrans` / `cn.CommitTrans`.

3. **`modDatabase.bas` (Serialización Canónica Inyectiva y Validación Preventiva):**
   - Serialización canónica por longitud prefijada (`Len(campo) & ":" & campo`), erradicando dependencias COM de .NET y riesgo de desbordamiento en 32 bits.
   - Validación preventiva de longitud de campos (`Habitacion <= 50`, `Motivo <= 255`) en `ValidarHojaBloqueosAntesDeGuardar`.

4. **`CopiarDatosHojasGijonSotoOviedo.bas` (Importación Atómica y Reversión Automática):**
   - Verificación previa de existencia de las 4 hojas involucradas (`wsBloqOrig`, `wsBloqDest`, `wsCalOrig`, `wsCalDest`).
   - Instantánea de memoria local previa (`arrBackupBloqDest`, `estadosBackupCalDest`) que se restaura automáticamente si ocurre cualquier excepción durante la copia.

5. **`ThisWorkbook.cls` (Manejo Interactivo de Colisiones):**
   - Manejo de excepciones en `Workbook_SheetDeactivate`, notificando al operador sobre cambios concurrentes y permitiendo la recarga sin pérdida de datos.
