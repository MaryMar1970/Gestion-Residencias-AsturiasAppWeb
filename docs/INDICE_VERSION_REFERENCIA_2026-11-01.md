# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: 2026-11-01 04:25 (Perfeccionamiento Quirúrgico Estricto de Borrado Parcial y Total — Concurrencia V5.2)

---

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  *Tamaño:* 6.809.197 bytes | *MD5:* `5727B66025C7278C95B5596CA22BD351` | *SHA-256:* `E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C`

- **Copia Timestamped en VersionReferencia (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm`
  *Tamaño:* 6.809.197 bytes | *MD5:* `5727B66025C7278C95B5596CA22BD351` | *SHA-256:* `E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C`

- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
  *Tamaño:* 6.746.112 bytes | *MD5:* `6237949FA8E9F503C9623F0A7294310D` | *SHA-256:* `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196`

- **Copia Timestamped Base de Datos (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-01_0425.accdb`
  *Tamaño:* 6.746.112 bytes | *MD5:* `6237949FA8E9F503C9623F0A7294310D` | *SHA-256:* `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196`

- **Copia Activa de Trabajo (Desarrollo y Pruebas):**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`
  *Tamaño:* 6.809.197 bytes | *MD5:* `5727B66025C7278C95B5596CA22BD351` | *SHA-256:* `E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C`

- **Copia Desplegada en Red (Producción Multiusuario Operativa):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`
  *Tamaño:* 6.809.197 bytes | *MD5:* `5727B66025C7278C95B5596CA22BD351` | *SHA-256:* `E3145EF12E49CAF22EEC07FEBB594A878A42320FB9A9EE5FC8F50860DBF61A6C`

---

### 2. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia estricta entre las intervenciones técnicas realizadas en la jornada 2026-11-01 y las copias de seguridad incrementales generadas en `H:\ResidenciaBD\BackupsDiarios\` y `H:\ResidenciaApp\bdas-multiusuario\`:

| Backup Incremental | Tamaño | Hash MD5 | Finalidad Técnica y Alcance Operativo |
| :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0230_PRE_OPTIMIZACION_BORRADO.xlsm` | 6.766.450 B | `4483E7AEB309A7DA1DE5408D0EB4B7DC` | **Línea Base Previa.** Estado anterior a las modificaciones de optimización de borrado. Reproduce la lentitud de ~1 s por fila y el defecto de filas vacías. |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0230_POST_OPTIMIZACION_BORRADO.xlsm` | 6.782.848 B | `53FEFA3D801E17163FEBD31F1C62A862` | **Optimización Inicial.** Incorporación de `Application.Union` y `rngEliminar.Delete Shift:=xlUp` en `BorradoSolicitudes.bas`, encolador por lotes en `ModuloSeguridadLocal.bas` y delegación matricial en `modDatabase.bas`. |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0345_VERSION_REFERENCIA.xlsm` | 6.792.119 B | `D616F3366D4AB18F507B5AE47A49DEE7` | **Consolidación Inicial de Referencia.** Primera versión estable de borrado instantáneo y compactación física de filas. |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0418_PRE_PERFECCIONAMIENTO_V5_2.xlsm` | 6.792.119 B | `D616F3366D4AB18F507B5AE47A49DEE7` | **Punto de Control Pre-Perfeccionamiento.** Respaldo de seguridad previo a la implantación estricta de la auditoría de concurrencia V5.2. |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0425_POST_PERFECCIONAMIENTO_V5_2.xlsm` | 6.809.197 B | `5727B66025C7278C95B5596CA22BD351` | **Perfeccionamiento V5.2 Implantado.** Incorpora monotonicidad de `idTransBase`, guarda de Lease Lock V5.2 (120 s), enrutamiento SICO-N2 en borrado total, discriminación de fórmulas pre-arrastradas y blindaje de eventos de aplicación. |
| `BDAS_Multiusuario_BACKUP_2026-11-01_0425_VERSION_REFERENCIA.xlsm` | 6.809.197 B | `5727B66025C7278C95B5596CA22BD351` | **Copia Oficial de Referencia Cierre de Sesión.** Homologada con 0 errores de compilación COM y 0 mojibake. |
| `Residencia_BE_BACKUP_2026-11-01_0425_VERSION_REFERENCIA.accdb` | 6.746.112 B | `6237949FA8E9F503C9623F0A7294310D` | **Respaldo Central de Base de Datos Access.** Estructura de tablas `Ordenes` y `Ordenes_Historico` 100% íntegra. |

- **Histórico de Referencias Previas en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0345.xlsm` (6.792.119 B | MD5: `D616F3366D4AB18F507B5AE47A49DEE7`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_2015.xlsm` (6.749.596 B | MD5: `6AE80126541D7D7F273B94929ECAE6A9`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-11-01_0345.accdb` (6.746.112 B | MD5: `6237949FA8E9F503C9623F0A7294310D`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-10-31_2015.accdb` (6.746.112 B | MD5: `6237949FA8E9F503C9623F0A7294310D`)

---

### 3. Inventario de Componentes y Mejoras Clave del Hito V5.2

1. **`ModuloSeguridadLocal.bas` (Encolador Masivo Offline SICO-N2):**
   - **Monotonicidad Estricta de ID_Transaccion y UUID:** Captura de `idTransBase` mediante `Application.WorksheetFunction.Max` antes del purgado de la cola local en RAM, garantizando correlatividad ascendente e impidiendo duplicidades de clave primaria en resincronizaciones.
   - **Purgado en Memoria RAM:** Una sola lectura masiva de `SYS_COLA_SYNC` a array Variant, filtrado en $O(1)$ con `Scripting.Dictionary`, acumulación mediante `Application.Union` y eliminación física en una única llamada nativa COM.
   - **Dimensionamiento a 13 Columnas:** Uso del enumerador nativo `colReintentos = 11`, `colUUID = 12`, `colCRC = 13` y cálculo de verificación por fila con `CalcularCRC16`.
   - **Un Solo Guardado a Disco:** Persistencia atómica final con `VolcarColaHaciaDiscoWAL`, reduciendo la latencia de 90 segundos a menos de 0,04 segundos.

2. **`modDatabase.bas` (Integración Transaccional y Control de Concurrencia V5.2):**
   - **Verificación Atómica de Lease Lock Activo ($120\text{ s}$):** Antes de eliminar órdenes en `EliminarOrdenesLoteBD` y `EliminarTodasOrdenesSedeBD`, consulta si alguna orden del lote está en edición activa por otro operador (`UsuarioModificacion <> tokenActual` y `DateDiff('s', UltimaModificacion, Now()) < 120`).
   - **Aborto Limpio y Formateador de Propietario:** Si detecta un bloqueo activo, cancela la transacción (`RollbackTrans`), cierra la conexión y notifica interactivamente mediante `FormatearPropietarioBloqueo` detallando el operador causante para evitar pisado de datos.
   - **Enrutamiento SICO-N2 en Borrado Total:** Soporte de contingencia offline en `EliminarTodasOrdenesSedeBD` mediante el parámetro opcional `colOrdenesOffline`, permitiendo delegar el lote completo a `EncolarBorradosLoteOffline` si no hay conectividad con Access.

3. **`BorradoDatosAlojamiento.bas` (Borrado Total Robusto de Alojamientos):**
   - **Desprotección y Desfiltrado Defensivo Previo:** Se retiran filtros automáticos (`ShowAllData`) y se desprotege la hoja antes de evaluar la última fila usada, evitando omitir filas ocultas por filtros activos.
   - **Discriminación de Fórmulas Pre-arrastradas:** Cálculo de `lastRowDatos` discriminando por hoja las columnas protegidas que contienen fórmulas continuas:
     - `RESIDENCIA GIJÓN`: Excluye cols 21, 28, 29.
     - `RESIDENCIA SOTO`: Excluye cols 17, 18, 21, 28, 29.
     - `RESIDENCIA OVIEDO`: Excluye cols 22, 32, 33.
   - **Recolección en RAM de Órdenes Activas:** Una sola lectura COM de la Columna A a array Variant para alimentar el enrutamiento SICO-N2.
   - **Aborto Transaccional del Vaciado Visual:** Si `EliminarTodasOrdenesSedeBD` devuelve `False`, la macro aborta la ejecución inmediatamente (`Exit Sub`) sin vaciar ninguna celda en Excel, preservando la coherencia front/back.

4. **`BorradoSolicitudes.bas` (Borrado Parcial Instantáneo y Compactación Atómica):**
   - **Eliminación Física Atómica con `rngEliminar.Delete Shift:=xlUp`:** Agrupación de todas las filas a borrar mediante `Application.Union` y supresión nativa en bloque, provocando que las filas inferiores asciendan instantáneamente sin dejar huecos vacíos.
   - **Preservación de Fórmulas Relativas:** Al ejecutarse un borrado nativo de filas completas de tabla, Excel reubica automáticamente los rangos de las fórmulas de importe y totales sin generar `#¡REF!`.
   - **Blindaje Estricto de Entorno:** Congelación incondicional de `Application.EnableEvents = False`, `ScreenUpdating = False` y `Calculation = xlCalculationManual` inmediatamente antes de la llamada de borrado, anulando cualquier ejecución parásita de eventos `Worksheet_Change`.

---

### 4. Garantías de Calidad y Gobernanza Técnica

- **Compilación VBA COM:** 100% verificada en instancia real de Excel COM; `VBE.CommandBars.FindControl(1, 578).Enabled == False` (**0 errores de compilación**).
- **Codificación Nativa:** Windows-1252 ANSI sin BOM en los 173 módulos exportados en `H:\ResidenciaApp\bdas-multiusuario\vba-modules\`.
- **Control Anti-Mojibake:** 0 caracteres corruptos detectados.
- **Regla Estricta de Toponimia:** Uso exclusivo del literal `GIJÓN` / `GIJON`. Prohibición total de escapes numéricos (`Chr(211)`).
- **Inmutabilidad del Golden Master:** La réplica `H:\ResidenciaBD\Plantilla\GoldenMaster\` no ha sido modificada.

