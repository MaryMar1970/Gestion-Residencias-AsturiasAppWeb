# RESIDENCIAAPP — ADENDA COMPLEMENTARIA UNIFICADA
## CIERRE DEL INFORME PREFLIGHT v1.0 Y CONCILIACIÓN DE EVIDENCIAS

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Documento base:** [`RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md`](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md)  
**Fase:** Cierre de la inspección técnica previa al laboratorio  
**Naturaleza:** Inspección documental y consultas de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  

---

## 1. IDENTIDAD Y RUTA FÍSICA DEL BACKEND Y VARIACIÓN DEL HASH

### 1.1. Inspección Documental de Conexiones DAO del PREFLIGHT

Se inspeccionaron estáticamente los scripts auxiliares utilizados en el PREFLIGHT para la consulta del motor DAO:
- [`test_dao.ps1`](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/scratch/test_dao.ps1)
- [`query_snapshot_full.ps1`](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/scratch/query_snapshot_full.ps1)
- [`detail_collisions_and_new.ps1`](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/scratch/detail_collisions_and_new.ps1)

**Resultados de la inspección:**
1. **Ruta física absoluta utilizada:** `H:\ResidenciaBD\Residencia_BE.accdb` (asignada de forma literal y unívoca en la variable `$dbPath`).
2. **Modo de apertura:** Estrictamente de solo lectura:
   ```powershell
   $db = $dao.OpenDatabase($dbPath, $false, $true) # Exclusive=False, ReadOnly=True
   ```
3. **Variables y parámetros:** No intervienen variables de entorno, cadenas de conexión dinámicas ni rutas alternativas.

---

### 1.2. Resolución de la Discrepancia Producción / Clon y Condición de Parada

Al cotejar la ruta utilizada con el inventario documental del proyecto, se acreditan los siguientes hechos objetivos:

| Archivo Físico | Ubicación | Tamaño (bytes) | Fecha Modificación | Rol según Documentación y Código |
| :--- | :--- | :--- | :--- | :--- |
| `Residencia_BE_TEST.accdb` | `H:\ResidenciaApp\bdas-multiusuario\transferencia_tecnica\laboratorio_pruebas\` | 6.729.728 | 22/10/2026 21:19:10 | Clon estático congelado para pruebas unitarias de Fase 5 (`INFORME_PRUEBAS_FASE5.md`). No contiene `BDAS_Historico_Snapshot`. |
| `Residencia_BE.accdb` | `H:\ResidenciaBD\` | 7.639.040 | 19/11/2026 22:42:09 | Base activa local (`DB_PATH_DEV` en `modDatabase.bas`). Designada como *«Backend de pruebas»* en las Órdenes 1R, 1S, 1T y 1U. |
| `Residencia_BE.accdb` | `H:\GIJ1\AYUDANTIA\01 RESIDENCIAS\BDAS Asturias\` | *(Red)* | *(Red)* | Base de red preproducción/terminales multiusuario (`DB_PATH_PREPROD` en `modDatabase.bas`). |

**Motivo de la discrepancia en la denominación:**
En los mandatos de las Órdenes 1R, 1S, 1T, 1U y PREFLIGHT emitidos por la Dirección Arquitectónica, se indicó textualmente como fuente autorizada:  
`- Backend de pruebas: H:\ResidenciaBD\Residencia_BE.accdb`.  
Por esta razón, los scripts de inspección ejecutaron sus lecturas sobre dicho archivo.

**Aplicación Estricta de la Condición de Parada:**  
Dado que la Dirección Arquitectónica clasifica formalmente `H:\ResidenciaBD\Residencia_BE.accdb` como la base de datos viva/producción y señala `transferencia_tecnica\laboratorio_pruebas\Residencia_BE_TEST.accdb` como el único clon de laboratorio:
1. Se constata que los scripts del PREFLIGHT consultaron la base de datos activa `H:\ResidenciaBD\Residencia_BE.accdb` (siempre en modo `ReadOnly = True`).
2. En cumplimiento estricto de la **Condición de Parada** del apartado 1.1: **se suspenden y bloquean de inmediato todas las aperturas y consultas contra `H:\ResidenciaBD\Residencia_BE.accdb` en esta orden**.
3. No se realiza ninguna nueva lectura sobre la base activa. Toda la información de conciliación requerida en los apartados siguientes se aporta a partir de los registros de ejecución completos y verificables ya obtenidos durante el PREFLIGHT, garantizando cero impacto sobre la base física.

---

### 1.3. Variación del Hash del Backend y Valoración de la Hipótesis

- **Hash anterior (Línea base Orden 1W):** `9617f99795b8266c2abb91b479bf64b9ece449ad15885ce40edbe62639dbbee9`.
- **Hash observado en PREFLIGHT:** `77db1d7ffeba4184451c0b9496dfa9d7eec27ffa3ae4c4a733686a042cbfdc03` (Marca temporal: `2026-11-19 22:42:09`, Tamaño: 7.639.040 bytes).

**Valoración de la hipótesis comunicada por la Dirección Arquitectónica:**
- La Dirección informó de que, con anterioridad al PREFLIGHT, realizó pruebas locales en Excel introduciendo datos en las columnas `REPARTIR COSTE` y `TAG`, reiniciando posteriormente la aplicación.
- En el código de `BDAS_Multiusuario.xlsm`, las acciones de inicio de sesión, apertura y cierre desencadenan escrituras automáticas en la tabla `LogActividad` vía `modDatabase.RegistrarAccion`. Además, cualquier conexión OLEDB con permisos de escritura modifica internamente las páginas de cabecera y el estado del archivo Access `.accdb`.
- Esta actividad manual coincide temporalmente con la marca del archivo (`22:42:09`), constituyendo una hipótesis técnica **altamente plausible**.
- Al haberse activado la condición de parada del apartado 1.1, no se han lanzado consultas adicionales sobre `LogActividad` en esta sesión, por lo que la explicación permanece categorizada rigurosamente como **plausible pero no demostrada mediante nueva consulta directa**.

---

## 2. CONCILIACIÓN EXHAUSTIVA DE FACTURAS DE `BDAS_HISTORICO_SNAPSHOT`

*(Datos recuperados íntegramente de las salidas de ejecución de `query_snapshot_full.ps1` y `detail_collisions_and_new.ps1`).*

La clave unívoca de consumo físico de facturación es:
$$\text{Clave} = (\text{Residencia}, \text{Year}(\text{FechaSalida}), \text{NumFactura})$$

### 2.1. Métricas Globales y Totales

| Métrica de Inspección | Valor Acreditado | Observaciones Metodológicas |
| :--- | :---: | :--- |
| **Total de filas físicas en la tabla** | **2.118** | Total de registros almacenados en `BDAS_Historico_Snapshot`. |
| **Filas físicas con `NumFactura > 0`** | **651** | Universo total de filas con factura emitida. |
| **Filas físicas con `NumFactura IS NULL`** | **1.467** | Expedientes archivados sin factura generada. |
| **Filas físicas con `NumFactura = 0`** | **0** | No existen registros con valor cero. |
| **Filas con `NumFactura > 0` y `FechaSalida IS NULL`** | **0** | El 100% de las facturas tiene fecha de salida válida. |
| **Rango de ejercicios observados** | **2025 – 2026** | Ejercicio mínimo: 2025 \| Ejercicio máximo: 2026. |
| **Claves únicas `(Residencia, Ejercicio, NumFactura)`** | **647** | Recuento estricto de claves primarias lógicas distintas. |
| **Diferencia (filas físicas vs. claves únicas)** | **4 claves (9 filas)** | Duplicados internos dentro del propio snapshot. |

---

### 2.2. Revisión de los Cuatro Duplicados Internos del Snapshot

Existen **4 claves de factura** que se repiten en el snapshot, agrupando a un total de **9 filas físicas**. Todas corresponden al ejercicio 2026:

| Residencia | Ejercicio | NumFactura | Filas Físicas | Referencias de Fila / Expediente (Anonimizadas) | Diagnóstico Técnico |
| :--- | :---: | :---: | :---: | :--- | :--- |
| **OVIEDO** | 2026 | **1** | **3** | SNAP_Fila_1 (Orden 108)<br>SNAP_Fila_2 (Orden 641)<br>SNAP_Fila_3 (Orden 2714) | **3 expedientes distintos** comparten la misma factura 1 en el snapshot (colisión de asignación histórica). |
| **OVIEDO** | 2026 | **3** | **2** | SNAP_Fila_4 (Orden 100)<br>SNAP_Fila_5 (Orden 2683) | **2 expedientes distintos** comparten la factura 3. |
| **OVIEDO** | 2026 | **4** | **2** | SNAP_Fila_6 (Orden 63)<br>SNAP_Fila_7 (Orden 2684) | **2 expedientes distintos** comparten la factura 4. |
| **SOTO** | 2026 | **3** | **2** | SNAP_Fila_8 (Orden 100)<br>SNAP_Fila_9 (Orden 101) | **2 expedientes distintos** comparten la factura 3. |

> [!IMPORTANT]
> Los 4 duplicados internos corresponden en todos los casos a **expedientes distintos** que recibieron indebidamente el mismo número físico de factura en el histórico, no a duplicaciones de volcado del mismo huésped.

---

### 2.3. Coincidencias Candidatas y Colisiones Potenciales con `Ordenes` (Activas)

Al cruzar las 651 filas del Snapshot con la tabla `Ordenes` activa mediante `(Residencia, Year(FechaSalida), NumFactura)`:
- **Coincidencias candidatas (mismo expediente: coinciden `NumOrden` y titular/DNI):** **634 tuplas**.
- **Colisiones relacionales (misma clave de factura pero distinto expediente):** **13 tuplas**.

#### Desglose Exhaustivo de las 13 Colisiones Potenciales:

| # | Residencia | Ejercicio | NumFactura | Expediente en Snapshot | Expediente en `Ordenes` Activa | Origen / Causa de la Colisión |
| :---: | :--- | :---: | :---: | :---: | :---: | :--- |
| **1** | OVIEDO | 2026 | 1 | Orden 108 | Orden 2714 | Duplicado interno Snapshot (Orden 2714 coincide; 108 colisiona). |
| **2** | OVIEDO | 2026 | 1 | Orden 641 | Orden 2714 | Duplicado interno Snapshot (Orden 2714 coincide; 641 colisiona). |
| **3** | OVIEDO | 2026 | 3 | Orden 2683 | Orden 100 | Duplicado interno Snapshot (Orden 100 coincide; 2683 colisiona). |
| **4** | OVIEDO | 2026 | 4 | Orden 2684 | Orden 63 | Duplicado interno Snapshot (Orden 63 coincide; 2684 colisiona). |
| **5** | OVIEDO | 2026 | 10 | Orden 2 | Orden 83 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **6** | OVIEDO | 2026 | 15 | Orden 2751 | Orden 109 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **7** | OVIEDO | 2026 | 16 | Orden 2750 | Orden 108 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **8** | OVIEDO | 2026 | 180 | Orden 183 | Orden 641 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **9** | OVIEDO | 2026 | 182 | Orden 175 | Orden 428 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **10** | OVIEDO | 2026 | 183 | Orden 171 | Orden 639 | Coincide con duplicidad histórica detectada en auditoría previa. |
| **11** | SOTO | 2026 | 1 | Orden 108 | Orden 109 | Desfase correlativo entre snapshot y activa. |
| **12** | SOTO | 2026 | 2 | Orden 103 | Orden 103 | Mismo `NumOrden` pero titular/DNI modificado entre versiones. |
| **13** | SOTO | 2026 | 3 | Orden 101 | Orden 100 | Duplicado interno Snapshot (Orden 100 coincide; 101 colisiona). |

> [!NOTE]
> Las 13 colisiones se desglosan en: 5 tuplas procedentes de los duplicados internos del Snapshot (OVIEDO 1, 3, 4 y SOTO 3), 6 tuplas que solapan exactamente con las anomalías históricas conocidas de OVIEDO (10, 15, 16, 180, 182, 183) y 2 casos de desincronización en SOTO (1 y 2).

---

### 2.4. Cruce Relacional con `Ordenes_Historico`

- **Coincidencias candidatas encontradas:** **0**.
- **Colisiones relacionales:** **0**.
- **Acreditación:** Ninguna de las 651 facturas del snapshot se encuentra en la tabla `Ordenes_Historico`.

---

### 2.5. Las Diez Claves Ausentes en Tablas Maestras

Existen **10 números de factura** presentes en `BDAS_Historico_Snapshot` que no existen ni en `Ordenes` ni en `Ordenes_Historico`:

| Residencia | Ejercicio | NumFactura | Filas en Snapshot | Estado Comprobado | Calificación Arquitectónica |
| :--- | :---: | :---: | :---: | :--- | :--- |
| **GIJÓN** | 2026 | **1** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **GIJÓN** | 2026 | **157** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **GIJÓN** | 2026 | **163** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **5** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **6** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **71** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **122** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **134** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **137** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |
| **OVIEDO** | 2026 | **292** | 1 | Ausente en `Ordenes` y en `Historico` | Consumo histórico huérfano. |

**Conclusión:** Cada una de estas 10 claves aparece exactamente en **1 sola fila física** del Snapshot (sin duplicados internos). Constituyen consumos documentales emitidos cuyos expedientes fueron purgados o nunca llegaron a sincronizarse en las tablas maestras.

---

### 2.6. Relación de Excepciones sin Resolver

1. **4 claves internas duplicadas en Snapshot (9 expedientes afectados):** OVIEDO (1, 3, 4) y SOTO (3).
2. **13 colisiones relacionales contra `Ordenes`:** Requieren resolución de precedencia antes de fijar el histórico.
3. **10 consumos huérfanos:** Deben reservarse en el ledger correlativo para que no se reemitan a nuevos huéspedes.

---

## 3. COPIAS HISTÓRICAS WAL `.BAK`

*(Evidencias extraídas de los archivos en `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\`)*.

### 3.1. Cuadro Comparativo de Archivos de Respaldo

| Parámetro | Copia 1: `SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_reintegrar_2026-11-02_0228` | Copia 2: `SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_revert_oviedo` |
| :--- | :--- | :--- |
| **Ruta física** | `C:\Users\jaime_ntymfau\AppData\Local\ResidenciaApp\SICO_N2\` | `C:\Users\jaime_ntymfau\AppData\Local\ResidenciaApp\SICO_N2\` |
| **Tamaño en bytes** | 265.175 bytes | 46.446 bytes |
| **Marca temporal** | `2026-11-02 02:18:21` | `2026-11-01 02:46:23` |
| **SHA-256** | `20cd431c01719022f6a2c89b53f8521c48c31059fed90cbbc48fb56cc07da119` | `892c19ce4f9c652151b77609dffd6768d98057d70ae518fe189633fcad1fa3e1` |
| **Líneas no vacías** | **2.133** | **377** |
| **Desglose de operaciones** | `LOG_OFFLINE`: 22<br>`ELIMINAR_ORDEN`: 2.111 | `LOG_OFFLINE`: 7<br>`ELIMINAR_ORDEN`: 370 |
| **Distribución de estados** | `PENDIENTE`: 2.133 (100%) | `PENDIENTE`: 377 (100%) |
| **Validación de CRC16** | **2.047 válidos** (incluye 22 con notación científica Excel).<br>86 discrepancias por formato legado. | **361 válidos** (incluye 5 con notación científica Excel).<br>16 discrepancias por formato legado. |
| **Naturaleza técnica** | Copia de seguridad previa a la reintegración masiva de Oviedo (02/11/2026). | Copia de seguridad previa a la reversión de pruebas de Oviedo (01/11/2026). |

### 3.2. Dictamen sobre la Cola WAL

1. Las dos copias `.wal.bak` son **instantáneas históricas de seguridad** creadas durante los incidentes de noviembre de 2026.
2. No constituyen colas activas ni deben ser reintegradas en la operativa ordinaria.
3. El archivo WAL activo `SYS_COLA_SYNC_SALÓN-PC.wal` **no existe en disco** porque el motor SICO-N2 lo elimina canónicamente (`Kill rutaWAL`) cuando la cola de transacciones pendientes está vacía (`ultFila < 2`).

---

## 4. RECTIFICACIÓN DOCUMENTAL DEL DOMINIO DE `TAG` Y `R`

Se rectifica formalmente cualquier mención previa que incluyera el cero dentro del dominio permitido.

La especificación técnica queda cerrada de forma definitiva:
- **Tipo de datos en Access:** `dbInteger` (DAO Type 3, entero con signo de 2 bytes).
- **Dominio admisible:** Enteros naturales estrictamente comprendidos entre **1 y 9999** ($1 \le x \le 9999$).
- **Valores prohibidos:**
  - Cero (`0`): **Inválido**.
  - Valores negativos ($< 0$): **Inválidos**.
  - Cadenas alfanuméricas y texto: **Inválidos**.
  - Valores con decimales: **Inválidos**.
- **Tratamiento del relleno:** Valores como `0001` y `1` son numéricamente idénticos; la base de datos almacena el entero escalar `1`.
- **Ausencia de valor:** Una celda vacía se persiste como `NULL`.
- **Requisito para Fase de Implementación:** La validación del rango 1–9999 deberá aplicarse en todas las puertas de entrada: cuadrícula Excel, procedimientos de pegado/edición VBA y reglas de validación relacionales en Access.

*(Los tres campos `REPARTIR COSTE` mantienen inalterado su dominio `NO`/`SI` y tipo `dbText(2)` nullable).*

---

## 5. LÍMITES DE ALCANCE Y ESTACIONES REMOTAS

1. **Estado de otras estaciones:** Durante el sondeo de red local, los demás puestos del entorno no resultaron accesibles vía SMB/UNC. Se mantiene la calificación formal de **NO INSPECCIONADAS** para todos los terminales remotos distintos de `SALÓN-PC`.
2. **Requisitos Obligatorios para el Procedimiento de Despliegue:**  
   Antes de sustituir el libro `.xlsm` por la versión v1.0 en cualquier estación cliente, el protocolo exigirá:
   - Iniciar la aplicación con conexión activa al backend oficial para permitir que el mecanismo FIFO procese y concilie cualquier operación diferida pendiente.
   - Constatar que la cola local no alberga transacciones en estado `PENDIENTE`, `SINCRONIZANDO`, `ERROR` o `CONFLICTO`.
   - Verificar la purga o preservación de copia del WAL local.
   - En caso de detectar operaciones no conciliadas, detener el despliegue en ese terminal y escalar la excepción.

---

## 6. CONFIRMACIÓN FINAL DE INTEGRIDAD Y DECISIÓN DE PUERTA

- **Declaración de Cero Modificaciones:** Durante la ejecución de esta adenda, el inspector técnico **NO ha modificado, creado, alterado ni compactado** ningún archivo `.xlsm`, base de datos `.accdb`, módulo VBA, tabla, registro ni archivo WAL.
- **Verificación de Git:** Comprobado mediante `git status --short`; el repositorio no presenta ninguna diferencia respecto al inicio de la sesión.
- **Control de Puerta Obligatorio:** Se ha verificado la regla estricta del literal `GIJÓN` y la ausencia total de mojibake o caracteres corruptos de codificación.

### Decisión de Puerta

En cumplimiento estricto de las directrices de la Dirección Arquitectónica:
- **Antigravity NO declara superada la Puerta 1.**
- **Antigravity NO inicia la Puerta 2 ni ejecuta ninguna acción de laboratorio.**
- La presente adenda probatoria se entrega a la Dirección Arquitectónica para que determine las pautas de cierre previas a la migración.