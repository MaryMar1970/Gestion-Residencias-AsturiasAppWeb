# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: 2026-11-03 00:15 (Cierre Oficial Jornada 02/11/2026: Resolución Dinámica Backend DB_PATH + Idempotencia EliminarOrdenBD y Saneamiento WAL SICO-N2 + Protección Hojas con UserInterfaceOnly en FechasPeticion y ListasDinamicasCPyPoblaciones + Coherencia Habitacional GIJÓN/OVIEDO)

---

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`  
  *Tamaño:* 6.616.947 bytes | *MD5:* `06E31C80A08574CA44F016850FBA6655` | *SHA-256:* `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`

- **Copia Timestamped en VersionReferencia (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-03_0015.xlsm`  
  *Tamaño:* 6.616.947 bytes | *MD5:* `06E31C80A08574CA44F016850FBA6655` | *SHA-256:* `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`

- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`  
  *Tamaño:* 6.754.304 bytes | *MD5:* `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | *SHA-256:* `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15`

- **Copia Timestamped Base de Datos (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-03_0015.accdb`  
  *Tamaño:* 6.754.304 bytes | *MD5:* `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | *SHA-256:* `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15`

- **Copia Activa de Trabajo (Desarrollo y Pruebas):**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6.616.947 bytes | *MD5:* `06E31C80A08574CA44F016850FBA6655` | *SHA-256:* `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`

- **Copia Desplegada en Red (Producción Multiusuario Operativa):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6.616.947 bytes | *MD5:* `06E31C80A08574CA44F016850FBA6655` | *SHA-256:* `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`

---

### 2. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia estricta entre las intervenciones técnicas realizadas en la jornada 2026-11-02 (cerrada a las 00:15 del 2026-11-03) y las copias de seguridad incrementales generadas en `H:\ResidenciaBD\BackupsDiarios\`:

| Backup Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Finalidad Técnica y Alcance Operativo |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.xlsm` | 6.430.806 B | `F73F5C628319611FD8C56CA773E60B67` | `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60` | **Línea Base Previa de Partida.** Estado anterior al hito de rutas dinámicas y validaciones de hoja. |
| `Residencia_BE_BACKUP_2026-11-02_0330_VERSION_REFERENCIA.accdb` | 6.746.112 B | `6237949FA8E9F503C9623F0A7294310D` | `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Respaldo previo de base de datos Access de partida. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2202_PRE_RUTAS.xlsm` | 6.379.116 B | `0C99D3A989A4F7B22B8255CCB575EBB8` | `2C1C4E9F95ABD46FA1941C9C4EF36A00314C24E11A9ADD2D98DEF6F90111E85B` | **Punto de Control Pre-Rutas.** Respaldo previo a sustituir rutas estáticas de backend. |
| `Residencia_BE_BACKUP_2026-11-02_2202_PRE_RUTAS.accdb` | 6.746.112 B | `6237949FA8E9F503C9623F0A7294310D` | `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Respaldo previo de base de datos Access. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2202_POST_RUTAS.xlsm` | 6.385.789 B | `549921707952BC6EAA423A6CBC107500` | `9239C28C7BD5D1CD76BFD282243430686BA6088A00C8BCE1E8FB2878C075448A` | **Rutas Dinámicas Implementadas.** Incorpora `Property Get DB_PATH` con resolución en cascada en `modDatabase.bas` y enlace en módulos satélite (`AsignarNumFactura`, `ModuloLOG`, `ModuloSeguridadLocal`). |
| `Residencia_BE_BACKUP_2026-11-02_2202_POST_RUTAS.accdb` | 6.746.112 B | `6237949FA8E9F503C9623F0A7294310D` | `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Respaldo posterior de base de datos Access. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2240_PRE_SANEAMIENTO_WAL.xlsm` | 6.385.789 B | `549921707952BC6EAA423A6CBC107500` | `9239C28C7BD5D1CD76BFD282243430686BA6088A00C8BCE1E8FB2878C075448A` | **Punto de Control Pre-Saneamiento WAL.** Respaldo antes de purgar operaciones residuales y corregir idempotencia. |
| `SYS_COLA_SYNC_PRE_SANEAMIENTO.wal.bak` | 220.155 B | — | — | Respaldo del archivo WAL físico `%LOCALAPPDATA%\...\SYS_COLA_SYNC_SALÓN-PC.wal` previo a la depuración. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2250_POST_SANEAMIENTO_WAL.xlsm` | 6.277.209 B | `BF866C8349F030DDF0246C8AC514C737` | `7AA6DAF21AD69CE410567A2901F6B905175D94092F680F4F191262D1BC87E551` | **Idempotencia y WAL Saneado.** Corrección de idempotencia en `EliminarOrdenBD` (`modDatabase.bas`), purga de 1.749 filas en `SYS_COLA_SYNC` y arranque en frío directo en `[CONECTADO]`. Reducción de >108 KB en el `.xlsm`. |
| `Residencia_BE_BACKUP_2026-11-02_2250_POST_SANEAMIENTO_WAL.accdb` | 6.746.112 B | `6DD968C7E5E84213266B255FB950FA44` | `F4E4D8FED4950590207C6772BB6DBAF8AF51DED8706CCFF607143CAB19F41B9F` | Respaldo de Access tras sincronización idempotente. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2310_PRE_FIX_PROTECCION_FECHAS.xlsm` | 6.277.209 B | `BF866C8349F030DDF0246C8AC514C737` | `7AA6DAF21AD69CE410567A2901F6B905175D94092F680F4F191262D1BC87E551` | **Pre-Fix Fechas.** Respaldo previo a subsanar el error de hoja protegida en Columna B (`FECHA PETICION`). |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2315_POST_FIX_PROTECCION_FECHAS.xlsm` | 6.286.240 B | `F3BC41E860FD5EE4D03BC76F037843D4` | `D55B4D34565128E70B2481A0A07D0341A6915C74450BA7FA46042AFD9F8AB3BE` | **Protección Fechas Corregida.** `FechasPeticion.bas` integrado con contraseña maestra de `ModuloConfigSegura` y restitución estricta de `UserInterfaceOnly:=True`. |
| `Residencia_BE_BACKUP_2026-11-02_2315_POST_FIX_PROTECCION_FECHAS.accdb` | 6.750.208 B | `877BB89415ACE72DC8F9FD417B67FAB5` | `DFA15AFDF7ABB18F9BB204B4F0E90A0E8A76154C66AA7E0799AEC919A5920759` | Respaldo de Access tras pruebas de fechas. |
| `BDAS_Multiusuario_BACKUP_2026-11-02_2340_POST_FIX_PROTECCION_LISTAS.xlsm` | 6.616.947 B | `06E31C80A08574CA44F016850FBA6655` | `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9` | **Protección Listas CP/Población Corregida.** Integración de `DesprotegerHojaSilenciosa` y `ReprotegerHojaSilenciosa` en `ListasDinamicasCPyPoblaciones.bas`, resolviendo la restricción de `.Validation` en hojas protegidas en `RESIDENCIA OVIEDO`, `GIJÓN` y `SOTO`. |
| `Residencia_BE_BACKUP_2026-11-02_2340_POST_FIX_PROTECCION_LISTAS.accdb` | 6.754.304 B | `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | Respaldo de Access tras pruebas de autocompletado CP y poblaciones. |
| `BDAS_Multiusuario_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.xlsm` | 6.616.947 B | `06E31C80A08574CA44F016850FBA6655` | `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9` | **Versión de Referencia Consolidada Final.** Copia oficial de cierre de jornada. |
| `Residencia_BE_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.accdb` | 6.754.304 B | `1A4836EAEBB42D5FAA3D6C06F8CDEBBA` | `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | Respaldo oficial de cierre de jornada de la base de datos Access. |

- **Histórico de Referencias Previas en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_2202.xlsm` (6.385.789 B | MD5: `549921707952BC6EAA423A6CBC107500`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm` (6.430.806 B | MD5: `F73F5C628319611FD8C56CA773E60B67`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0110.xlsm` (6.838.431 B | MD5: `33AC497F6CF682AFCB79DDAAA07446C2`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm` (6.809.197 B | MD5: `5727B66025C7278C95B5596CA22BD351`)

---

### 3. Inventario de Componentes y Mejoras Clave del Hito 2026-11-03 00:15

1. **`modDatabase.bas` (Resolución Dinámica de Backend y Caché de Ruta):**
   - Implementación de `Property Get DB_PATH`: algoritmo de cascada con memoria caché (`m_sRutaCacheBD`) que resuelve dinámicamente en orden:
     1. Directorio del libro activo (`ThisWorkbook.Path & "\Residencia_BE.accdb"`).
     2. Red de terminales preproducción (`H:\GIJ1\AYUDANTIA\01 RESIDENCIAS\BDAS Asturias\Residencia_BE.accdb`).
     3. Entorno de desarrollo local (`H:\ResidenciaBD\Residencia_BE.accdb`).
   - Implementación de `LimpiarCacheRutaBD` para invalidación controlada.
   - Reemplazo y desacoplamiento de constantes fijas en módulos satélite (`AsignarNumFactura`, `ModuloLOG`, `ModuloSeguridadLocal`).

2. **`modDatabase.bas` (Principio de Idempotencia en `EliminarOrdenBD`):**
   - Corrección matemática del flujo de eliminación: cuando una orden ya ha sido borrada previamente (`rsCheck.EOF = True`), la operación devuelve `EliminarOrdenBD = True` y ejecuta `CommitTrans` en lugar de `False` y `RollbackTrans`.
   - Se elimina el bucle infinito de reintentos fallidos en la sincronización diferida SICO-N2.

3. **Búfer WAL y Hoja Oculta `SYS_COLA_SYNC` (Saneamiento Integral):**
   - Purga física de 1.749 filas residuales acumuladas en la hoja oculta `SYS_COLA_SYNC`.
   - Eliminación del archivo desfasado de 220 KB en `%LOCALAPPDATA%\...\SYS_COLA_SYNC_SALÓN-PC.wal` (previamente respaldado).
   - `ContarCambiosPendientes()` pasa a **0**, reduciendo el tamaño del libro en más de 108 KB e iniciando directamente en `[CONECTADO]`.

4. **`FechasPeticion.bas` (Protección Segura de Hojas con `UserInterfaceOnly`):**
   - Erradicación de contraseñas vacías hardcodeadas (`hoja.Unprotect password:=""`).
   - Enlace directo con `ModuloConfigSegura.ObtenerPasswordHojas()`.
   - Reprotección atómica en `CleanUp:` y `ErrHandler:` preservando invariablemente `UserInterfaceOnly:=True` y todos los permisos de edición (`AllowFormattingCells`, `AllowSorting`, etc.).

5. **`ListasDinamicasCPyPoblaciones.bas` (Gestión Silenciosa de Validaciones en Hojas Protegidas):**
   - Incorporación de `DesprotegerHojaSilenciosa` y `ReprotegerHojaSilenciosa` consumiendo la contraseña maestra unificada.
   - Envoltura atómica de `LimpiarValidacionSilenciosa`, `CrearValidacionListaSilenciosa` y `CrearValidacionAutocompletadoOptimizada`.
   - Superación de la restricción nativa de Excel sobre la colección `.Validation` en celdas con `Locked = False` en hojas protegidas.

6. **`Hoja2.cls` (`RESIDENCIA GIJÓN`) y `Hoja24.cls` (`RESIDENCIA OVIEDO`) (Coherencia Habitacional):**
   - Desacoplamiento del mapeo de columnas para validación de PAX vs. Habitaciones Facturadas y Adjudicadas:
     * En GIJÓN: Columna S mapeada como *Habitaciones Adjudicadas*, verificando $Q + R = S$.
     * En OVIEDO: Columna S mapeada como *Camas Supletorias* (supeditadas a $R > 0$) y Columna T como *Habitaciones Adjudicadas*.

---

### 4. Control de Calidad ECRMC y Protocolo de Codificación

- **Compilación VBE:** Verificación ejecutada vía automatización COM Excel (`FindControl(1, 578)`): **0 errores de compilación**.
- **Control de Mojibake:** Escaneo automatizado en Python sobre la totalidad de los archivos de código: **0 caracteres corruptos** (100% nativo **ANSI Windows-1252**).
- **Regla Estricta de Sede:** Cumplimiento del 100% en el uso del literal **`GIJÓN`** (0 violaciones de `Chr(211)` o concatenaciones prohibidas).

