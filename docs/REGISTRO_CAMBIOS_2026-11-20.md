# REGISTRO DE CAMBIOS Y ACTIVIDAD
## Fecha: 2026-11-20 | Hito: Cierre de Sesión — Auditoría de Evidencias Arquitectónicas y Preflight v1.0
### Ámbito: Consolidación de Informes de Inspección (Órdenes 1Q a 1X, Preflight y Adenda Unificada), Respaldo Incremental de Seguridad y Retención en Puerta 1

---

### 1. Resumen de la Intervención y Contexto Técnico

En la jornada de trabajo del **20 de Noviembre de 2026**, se completó con éxito el ciclo integral de inspecciones técnicas, auditorías estáticas de código, extracción de evidencias y comprobaciones complementarias de preflight encomendadas por la Dirección Arquitectónica para cerrar el diseño de la arquitectura v1.0 (identidad persistente mediante `IdInstanciaOrden`, conciliación transaccional de facturación y saneamiento de sincronización offline SICO-N2):

1. **Orden 1Q y 1Q-A (Revisión de Reglas de Ejercicio y Numeración):** Determinación determinista del ejercicio de orden por `Year(FechaPeticion)` y de factura por `Year(FechaSalida)`. Análisis de correspondencia física Excel-Access y unicidad `(Residencia, Anio, Numero)`.
2. **Orden 1R (Inspección de Fechas y Trazabilidad Histórica de Facturas):** Detección de expedientes sin fechas válidas y análisis forense del comportamiento de sustitución de versiones en `Ordenes_Historico`.
3. **Orden 1S (Identidad y Dependencias de `Id` y `BDAS_Historico_Snapshot`):** Auditoría del campo autonumérico `Id`, trazabilidad de orígenes de `BDAS_Historico_Snapshot` (contingencia de Oviedo de noviembre de 2026) y viabilidad técnica de GUID de 128 bits offline.
4. **Orden 1T (Extracción de Persistencia Offline, Mapeo Excel-Access y `ContadorFacturas`):** Mapeo exhaustivo de 48 columnas, verificación del tipo físico `ContadorFacturas.Ejercicio` (`dbInteger`, 2 bytes) y auditoría de los 5 campos de usuario sin persistencia previa (`REPARTIR COSTE` x3, `TAG` y `R`).
5. **Orden 1U (Verificación de Integridad y Resolución de Discrepancias):** Resolución unívoca de metadatos, tamaños y hashes de los componentes autorizados.
6. **Orden 1V (Persistencia Local de Identidad y Mecánica de Altas/Modificaciones):** Determinación de la persistencia de `IdInstanciaOrden` en columna técnica de Excel y trazabilidad de todos los flujos de alta en el sistema.
7. **Orden 1W (Verificación Puntual de Procedimientos Críticos):** Análisis estático de `EjecutarUpdateCondicionalOrden`, `EliminarOrdenBD` y `RestaurarOrdenDesdeHistorico`.
8. **Orden 1X (Rangos Operativos, Desplazamiento, Inserción, Ordenación y Limpieza):** Verificación de la cobertura de la columna candidata `AO` (columna 41) para albergar `IdInstanciaOrden` en las hojas de GIJÓN, OVIEDO y SOTO.
9. **Informe PREFLIGHT v1.0:** Auditoría de partida previa al laboratorio sobre los cinco campos huérfanos, facturas del snapshot y estado de `SYS_COLA_SYNC` y WAL.
10. **Adenda Complementaria Unificada PREFLIGHT v1.0:** Resolución de la discrepancia de identidad del backend (acreditando que las lecturas previas fueron en `H:\ResidenciaBD\Residencia_BE.accdb` en modo `ReadOnly = True`), activación estricta de la condición de parada de solo lectura, conciliación rigurosa de las facturas del snapshot (651 filas, 647 claves únicas, 4 duplicados internos, 13 colisiones y 10 huérfanas), inspección de los respaldos históricos `.wal.bak`, rectificación documental del dominio de `TAG` y `R` (1 a 9999) y retención en Puerta 1.

---

### 2. Inventario de Informes Generados y Preservados en `docs/`

Todos los informes elaborados durante la sesión han sido consolidados y copiados en [`H:\ResidenciaApp\docs\`](file:///H:/ResidenciaApp/docs/):

1. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1Q.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1Q.md)
2. [`RESIDENCIAAPP_INFORME_COMPLEMENTARIO_ORDEN_1Q_A.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_COMPLEMENTARIO_ORDEN_1Q_A.md)
3. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1R.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1R.md)
4. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md)
5. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1T.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1T.md)
6. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1U.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1U.md)
7. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1V.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1V.md)
8. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1W.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1W.md)
9. [`RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1X.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1X.md)
10. [`RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md)
11. [`RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md`](file:///H:/ResidenciaApp/docs/RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md)

---

### 3. Trazabilidad Estricta: Copias de Seguridad Incrementales al Cierre

En cumplimiento del protocolo de seguridad del proyecto, al término de la sesión se han generado los respaldos de seguridad correspondientes en [`H:\ResidenciaBD\BackupsDiarios\`](file:///H:/ResidenciaBD/BackupsDiarios/):

| Archivo de Respaldo Incremental | Tamaño | Hash SHA-256 | Estado y Alcance |
| :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-20_0045_CIERRE_SESION_EVIDENCIAS_v1_0.xlsm` | 6.621.963 B | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **Libro Activo Intacto.** 100% coincidente con la Versión de Referencia Oficial (cero modificaciones). |
| `Residencia_BE_BACKUP_2026-11-20_0045_CIERRE_SESION_EVIDENCIAS_v1_0.accdb` | 7.639.040 B | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **Base de Datos Activa.** Estado intacto tras inspecciones de solo lectura. |

---

### 4. Sesión de Tarde: Ejecución de Puerta 2 (Laboratorio Técnico — Paso 1 y Adenda de Seguridad/WAL)

1. **Apertura de Puerta 2:** Recibida autorización formal para ejecutar el **Paso 1** (Creación del entorno de laboratorio, clonación verificada y aislamiento de rutas).
2. **Entorno de Laboratorio Aislado:**
   - Creado en `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\` con subcarpetas `BASELINE\`, `TRABAJO\`, `EVIDENCIAS\` y `TRABAJO\WAL_TEST\`.
   - Control de integridad SHA-256 verificado: `Residencia_BE.accdb` (`77DB1D7F...`) y `BDAS_Multiusuario.xlsm` (`B9117659...`) idénticos en origen, `BASELINE` y `TRABAJO`.
3. **Cierre de Comprobaciones Complementarias (Adenda de Seguridad y WAL):**
   - **Permisos NTFS:** `BASELINE` restringido a solo lectura `(OI)(CI)(RX)` eliminando grupos no autorizados; inmutabilidad verificada empíricamente con `PermissionError` ante intento de escritura. `TRABAJO` restringido a la cuenta autorizada.
   - **Conexión Real 32 bits:** Apertura exitosa de DAO (`DAO.DBEngine.120`) y ADODB (`Microsoft.ACE.OLEDB.12.0`) en 32 bits contra `TRABAJO\Residencia_BE.accdb` (2.127 órdenes y 291 tarifas leídas). Comprobado bloqueo local temporal `TRABAJO\Residencia_BE.laccdb` y cero conexiones a producción (`H:\ResidenciaBD`). Cero `.laccdb` residuales tras cierre.
   - **Aclaración COM/VBS:** Explicada la causa por la que `xl.Run("DB_PATH")` arrojó error (es `Property Get`) y la idoneidad de `GetConnectionString()` (que devolvió con código 0 la ruta aislada).
   - **Aislamiento Obligatorio SICO-N2/WAL:** Adaptada quirúrgicamente **exclusivamente la copia de trabajo** (`TRABAJO\BDAS_Multiusuario.xlsm`) para dirigir el WAL a `TRABAJO\WAL_TEST\SYS_COLA_SYNC_<COMPUTERNAME>.wal` con soporte multi-escenario (`SICO_WAL_SUBDIR`). Verificada la resolución en runtime mediante Excel COM 32 bits. El perfil real `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\` permanece 100% intacto con sus 3 archivos históricos originales.
4. **Informes Generados:**
   - [`INFORME_PASO1_SETUP_LABORATORIO.md`](file:///H:/ResidenciaApp/docs/INFORME_PASO1_SETUP_LABORATORIO.md)
   - [`INFORME_ADENDA_PASO1_SEGURIDAD_Y_WAL.md`](file:///H:/ResidenciaApp/docs/INFORME_ADENDA_PASO1_SEGURIDAD_Y_WAL.md)
5. **Condición de Parada Activa:**
   - Estado retenido a la espera de la revisión y conformidad formal de la Dirección Arquitectónica.
   - Ninguna prueba DDL, DML ni funcional ha sido ejecutada.
   - Cero procesos Excel o Access en ejecución; cero ficheros de bloqueo `.laccdb`.
   - Regla `GIJÓN`: Cumplida al 100%. Cero Mojibake.