# RESIDENCIAAPP — INFORME DE INSPECCIÓN TÉCNICA FINAL
## PREFLIGHT v1.0 — COMPROBACIONES DE ESPECIFICACIÓN ARQUITECTÓNICA

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Cierre del diseño de la especificación arquitectónica v1.0  
**Naturaleza:** Inspección estática, OpenXML y consultas DAO de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha de emisión:** 19 de noviembre de 2026  
**Documento de referencia:** Apartado 12 de `ResidenciaApp_Especificacion_Arquitectonica_v1_0.md`  

---

## 1. PREPARACIÓN, TRAZABILIDAD Y CONTROL DE FUENTES

En estricto cumplimiento de `AGENTS.md` y `PROTOCOLO_CAMBIOS.md`, la presente auditoría técnica se ha ejecutado de manera pasiva y no destructiva:
- No se ha abierto el libro Excel mediante interfaces que disparen macros ni eventos automáticos (`Worksheet_Activate`, `Worksheet_Change`, `Workbook_Open`).
- La inspección de las hojas de Excel se ha realizado mediante análisis directo del paquete OpenXML comprimido.
- Las consultas a la base de datos se han ejecutado sobre el **backend de pruebas autorizado** (`H:\ResidenciaBD\Residencia_BE.accdb`) utilizando llamadas DAO en modo explícito de solo lectura (`ReadOnly = True`), preservando completamente intactas las réplicas maestras de producción en `Plantilla\GoldenMaster` y `Plantilla\VersionReferencia`.
- No se emite en este informe ningún dato de carácter personal de huéspedes (nombres, DNI, teléfonos, direcciones ni identificadores de clientes), en riguroso cumplimiento del marco RGPD.

### 1.1. Integridad de los Archivos Objetivo Inspeccionados

| Componente | Ruta Absoluta | Tamaño (Bytes) | Fecha de Modificación | SHA-256 |
| :--- | :--- | :---: | :---: | :--- |
| **Libro Activo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | 2026-11-09 20:32:02 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` |
| **Backend Pruebas** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7.639.040 | 2026-11-19 22:42:09 | `77db1d7ffeba4184451c0b9496dfa9d7eec27ffa3ae4c4a733686a042cbfdc03` |
| **Módulo Seguridad**| `H:\ResidenciaApp\bdas-multiusuario\vba-modules\ModuloSeguridadLocal.bas` | 51.523 | 2026-11-09 19:42:06 | `7486f0ff528c11e3b5e40e6983792cb0d09995c739bf824bf28cf0810aa157ff` |
| **Módulo Database** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 298.794 | 2026-11-10 15:45:00 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` |

---

## 2. COMPROBACIÓN A — COMPATIBILIDAD DE LAS VALIDACIONES DE LOS CINCO CAMPOS

Se inspeccionó la estructura OpenXML de las hojas operativas (`sheet2.xml` para GIJÓN, `sheet23.xml` para SOTO y `sheet34.xml` para OVIEDO), los estilos del libro (`xl/styles.xml`) y el código VBA correspondiente, verificando las restricciones de datos, formatos de celda y consumidores para los cinco campos funcionales.

### 2.1. Resultados de la Inspección OpenXML

| Hoja | Columna | Campo | Validación OpenXML (`dataValidation`) | Formato de Celda (`numFmt`) | Fórmulas en Celdas | Estado de Celdas Guardadas |
| :--- | :---: | :--- | :--- | :---: | :---: | :--- |
| **RESIDENCIA GIJÓN** | `AD` (30) | `REPARTIR COSTE` | `type="list"`, `formula1="""NO,SI"""`<br>`sqref="AD2:AD2663"`, `allowBlank="1"` | `General` | Ninguna (0) | Vacías (`with_values = 0`). Resultado esperado. |
| **RESIDENCIA SOTO** | `AD` (30) | `REPARTIR COSTE` | `type="list"`, `formula1="""NO,SI"""`<br>`sqref="AD2:AD1048576"`, `allowBlank="1"` | `General` | Ninguna (0) | Vacías (`with_values = 0`). Resultado esperado. |
| **RESIDENCIA OVIEDO** | `AC` (29) | `REPARTIR COSTE` | `type="list"`, `formula1="""NO,SI"""`<br>`sqref="AC2:AC1724"`, `allowBlank="1"` | `General` / `0` | Ninguna (0) | Vacías (`with_values = 0`). Resultado esperado. |
| **RESIDENCIA OVIEDO** | `AD` (30) | `TAG` | Sin validación de lista (entrada numérica libre) | `General` | Ninguna (0) | Vacías (`with_values = 0`). Resultado esperado. |
| **RESIDENCIA OVIEDO** | `AE` (31) | `R` | Sin validación de lista (entrada numérica libre) | `General` | Ninguna (0) | Vacías (`with_values = 0`). Resultado esperado. |

### 2.2. Acreditaciones Técnicas de Compatibilidad

1. **Compatibilidad de `REPARTIR COSTE`:**
   - La regla de validación de Excel en las tres residencias impone de forma estricta la lista `"NO,SI"`.
   - No existe contradicción alguna con el dominio funcional confirmado (`NO` / `SI`).
   - El tipo de campo físico en Access `dbText(2)` nullable se alinea con la longitud exacta requerida (máximo 2 caracteres ANSI).
2. **Compatibilidad de `TAG` y `R`:**
   - En `RESIDENCIA OVIEDO`, las columnas `AD` (30) y `AE` (31) están configuradas con formato `General`. No tienen máscaras de texto forzado (`@`) ni formatos decimales (`0.00`).
   - No contienen fórmulas calculadas ni referencias dependientes.
   - En el código VBA, ningún procedimiento trata estas columnas como cadenas alfanuméricas complejas ni realiza cálculos de coma flotante sobre ellas.
   - El tipo físico Access propuesto, **`dbInteger` (DAO Type 3, entero de 2 bytes con signo, rango $-32.768$ a $32.767$)**, es idóneo y cubre con holgura el dominio de enteros de hasta 4 dígitos ($0 \le \text{valor} \le 9.999$). El valor vacío en la celda se corresponderá de forma natural con `NULL` en base de datos.
3. **Consumidores VBA Adicionales Verificados:**
   - En [Hoja24.cls#L549](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja24.cls#L549), la columna `AC` (29) participa en el refresco de `ValorAnteriorOviedo` para listas desplegables.
   - En [Hoja2.cls#L103](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja2.cls#L103) y [Hoja20.cls#L512](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja20.cls#L512), la columna 30 (`AD`) está incluida en el rango admitido de cambios y refresco de valor previo.
   - Las columnas `AD` y `AE` de Oviedo no son consumidas por reglas de cálculo de facturación ni calendarios; operan exclusivamente como atributos directos del expediente.
4. **Constatación de Celdas Vacías en el Archivo Guardado:**
   - En las tres hojas operativas del libro `BDAS_Multiusuario.xlsm`, las celdas de las filas de datos para estas 5 columnas están vacías.
   - Tal como indicó la Dirección Arquitectónica, **este es el resultado esperado** dado que dichos campos nunca tuvieron persistencia relacional en Access. **No constituye un bloqueo de migración ni requiere backfill de datos**.

---

## 3. COMPROBACIÓN B — FACTURAS POSITIVAS EN `BDAS_Historico_Snapshot`

Mediante consultas directas DAO en modo de solo lectura sobre la tabla `BDAS_Historico_Snapshot` del backend de pruebas (`Residencia_BE.accdb`), se extrajeron las métricas relacionales agregadas requeridas.

### 3.1. Métricas Generales del Snapshot

| Métrica Solicitada | Valor Acreditado | Observaciones Técnicas |
| :--- | :---: | :--- |
| **Total de filas en `BDAS_Historico_Snapshot`** | **2.118** | Volumen total de registros históricos en la tabla snapshot. |
| **1. Recuento con `NumFactura > 0`** | **651** | Expedientes con factura formalmente asignada. |
| **7. Recuento con `NumFactura = 0`** | **0** | No existen registros con factura cero en el snapshot. |
| **Excepción: `NumFactura IS NULL`** | **1.467** | Expedientes sin facturar o solicitudes no liquidadas ($651 + 1.467 = 2.118$). |
| **4. Recuento `NumFactura > 0` con `FechaSalida IS NULL`** | **0** | El 100% de las facturas positivas disponen de `FechaSalida` válida. |
| **3. Ejercicio mínimo observado** | **2025** | Fecha más antigua entre las facturas positivas. |
| **3. Ejercicio máximo observado** | **2026** | Ejercicio vigente en las facturas del snapshot. |

### 3.2. Distribución Agregada por Residencia y Ejercicio

| Residencia | Ejercicio (`Year(FechaSalida)`) | Total Facturas Positivas | Rango Mínimo .. Máximo de Factura |
| :--- | :---: | :---: | :---: |
| **GIJÓN** | 2026 | 30 | 1 .. 179 |
| **OVIEDO** | 2025 | 13 | 994 .. 1023 |
| **OVIEDO** | 2026 | 593 | 1 .. 1029 |
| **SOTO** | 2026 | 15 | 1 .. 14 |
| **TOTAL** | — | **651** | — |

### 3.3. Duplicados Internos en el Snapshot por Clave Completa
Consulta ejecutada sobre la clave `(Residencia, Year(FechaSalida), NumFactura)`:
- **Total de claves con colisión interna:** **4 claves** (que agrupan un total de 9 filas físicas en el snapshot).

| Residencia | Ejercicio | Nº Factura | Repeticiones en Snapshot | Naturaleza del Registro |
| :--- | :---: | :---: | :---: | :--- |
| **OVIEDO** | 2026 | 1 | 3 | Mismo número de factura asignado a 3 filas en el snapshot. |
| **OVIEDO** | 2026 | 3 | 2 | Mismo número de factura asignado a 2 filas en el snapshot. |
| **OVIEDO** | 2026 | 4 | 2 | Mismo número de factura asignado a 2 filas en el snapshot. |
| **SOTO** | 2026 | 3 | 2 | Mismo número de factura asignado a 2 filas en el snapshot. |

### 3.4. Cruce de Tuplas Positivas del Snapshot con las Tablas de Expedientes

Se realizó el cruce relacional entre las tuplas `(Residencia, Ejercicio, NumFactura)` del snapshot y las tablas maestras `Ordenes` y `Ordenes_Historico`:

```sql
SELECT s.Residencia, Year(s.FechaSalida), s.NumFactura ...
FROM BDAS_Historico_Snapshot s
INNER JOIN Ordenes o ON s.Residencia = o.Residencia 
   AND Year(s.FechaSalida) = Year(o.FechaSalida) 
   AND s.NumFactura = o.NumFactura
```

#### Resultados del Cruce:
1. **Contra la tabla `Ordenes` (Activas):**
   - **Coincidencias Exactas (mismo expediente: coinciden `NumOrden` y `DNI`):** **634 tuplas**.
   - **Colisiones Potenciales (misma factura y año, pero distinto expediente/huésped):** **13 tuplas**.
     - Detalle de colisiones en OVIEDO (2026): Facturas 1, 3, 4, 10, 15, 16, 180, 182, 183.
     - Detalle de colisiones en SOTO (2026): Facturas 1, 2, 3.
2. **Contra la tabla `Ordenes_Historico`:**
   - **Coincidencias Exactas:** **0**.
   - **Colisiones Potenciales:** **0**.
3. **Tuplas Positivas de Snapshot Ausentes en Ambas Tablas:**
   - Existen **10 números de factura distintos** presentes en `BDAS_Historico_Snapshot` que no constan actualmente ni en `Ordenes` ni en `Ordenes_Historico`:
     - **GIJÓN (2026):** Facturas **1**, **157**, **163** (3 facturas).
     - **OVIEDO (2026):** Facturas **5**, **6**, **71**, **122**, **134**, **137**, **292** (7 facturas).

> [!NOTE]
> **Implicación para el Backfill de `RegistroNumerosFactura`:**  
> Estas 10 facturas huérfanas representan consumos históricos reales que no deben ser reemitidos. Al incorporar estas 10 claves al seed/backfill de la tabla de facturación correlativa, se previene que una nueva orden genere una factura ya existente documentalmente.

---

## 4. COMPROBACIÓN C — ESTADO ACTUAL DE `SYS_COLA_SYNC` Y WAL

### 4.1. Hoja Técnica `SYS_COLA_SYNC` (Libro `BDAS_Multiusuario.xlsm`)

Se inspeccionó la hoja técnica oculta `SYS_COLA_SYNC` (`sheet49.xml`) directamente desde el OpenXML del libro activo:
- **Cabeceras detectadas (Fila 1):** `ID_Transaccion`, `FechaHoraLocal`, `Usuario`, `Residencia`, `NumOrden`, `TipoOperacion`, `CampoModificado`, `ValorAnterior`, `ValorNuevo`, `EstadoSync`, `Reintentos`, `UUID_Idempotencia`, `CRC_Integridad`.
- **Total de filas de transacción con datos útiles:** **0 FILAS (COLA TÉCNICA VACÍA)**.
- **Análisis de filas residuales:** El archivo XML conserva 1.748 filas etiquetadas estructuralmente (filas 2 a 1.749) que contienen exclusivamente punteros de estilo de celda vacíos (`s=405`, `s=1046`). Todas las celdas de las columnas `A` a `M` tienen contenido `v=None` / nulo.
- **Métricas de la Cola en Libro:**
  - Filas en `PENDIENTE`: **0**.
  - Filas en `SINCRONIZANDO`: **0**.
  - Filas en `ERROR`: **0**.
  - Filas en conflicto: **0**.
  - Registros de CRC válidos / inválidos: **0** (sin transacciones activas).
  - UUIDs vacíos o repetidos: **0**.
- **Conclusión de la Hoja Técnica:** La cuadrícula técnica de contingencia no alberga ninguna operación diferida pendiente de conciliar.

---

### 4.2. Archivos WAL Locales en Disco

La arquitectura SICO-N2 establece que el archivo WAL reside en el perfil local de cada estación:
`%LOCALAPPDATA%\ResidenciaApp\SICO_N2\SYS_COLA_SYNC_<COMPUTERNAME>.wal`

#### A. Inspección en la Estación Activa (`SALÓN-PC`)
- **Directorio:** `C:\Users\jaime_ntymfau\AppData\Local\ResidenciaApp\SICO_N2\`
- **Archivo Activo `SYS_COLA_SYNC_SALÓN-PC.wal`:** **AUSENTE / NO EXISTE**.
  - **Acreditación del motivo:** El código oficial en [ModuloSeguridadLocal.bas#L1468](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1468) y [L1494](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1494) establece explícitamente:
    ```vba
    If ultFila < 2 Then If Len(Dir$(rutaWAL)) > 0 Then Kill rutaWAL
    ...
    If escritos = 0 Then If Len(Dir$(rutaWAL)) > 0 Then Kill rutaWAL
    ```
    Al no existir transacciones pendientes ni encoladas, el archivo físico `.wal` es eliminado de forma intencional por el propio motor al completarse el vaciado.
- **Archivos Históricos de Respaldo Localizados:**
  1. `SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_reintegrar_2026-11-02_0228` (265.175 bytes, modificado 2026-11-02 02:18, SHA-256: `20cd431c01719022f6a2c89b53f8521c48c31059fed90cbbc48fb56cc07da119`).
  2. `SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_revert_oviedo` (46.446 bytes, modificado 2026-11-01 02:46, SHA-256: `892c19ce4f9c652151b77609dffd6768d98057d70ae518fe189633fcad1fa3e1`).
  *(Corresponden a instantáneas de seguridad creadas durante las pruebas de integración históricas del 1 y 2 de noviembre de 2026; no son archivos WAL activos).*

#### B. Estaciones Remotas del Entorno Multiusuario
- Se ejecutó el sondeo de red local (`net use`, `net view`). No existen recursos compartidos ni puestos de trabajo remotos alcanzables mediante rutas UNC desde este equipo.
- **Declaración expresa:** **Las demás posibles estaciones clientes del sistema quedan clasificadas como NO INSPECCIONADAS** debido a su aislamiento físico/de red respecto a esta consola.

---

### 4.3. Dictamen de Transición Arquitectónica

1. **Estado en el Equipo Local (`SALÓN-PC`):**
   Tanto la hoja técnica `SYS_COLA_SYNC` en el libro maestro como el archivo `.wal` local en disco están **completamente vacíos y sin operaciones pendientes**.
2. **Viabilidad de Transición:**
   Para esta estación y el archivo central del proyecto, **el cambio de formato y esquema hacia v1.0 (incorporando `IdInstanciaOrden` y las nuevas tablas) puede prepararse de forma directa sin requerir procesos de drenaje ni migración de colas pendientes**, manteniendo el protocolo de copias de seguridad de rigor.
3. **Recomendación Preventiva para el Despliegue Global:**
   Dado que no es posible auditar de forma remota otras máquinas físicas que pudieran haber operado en offline, el protocolo de implantación v1.0 deberá exigir que cualquier otro equipo cliente inicie sesión conectado a la base de datos antes de sustituir el libro `.xlsm`, garantizando que cualquier cola local preexistente quede conciliada antes de actualizar la versión.

---

## 5. TABLA SÍNTESIS DE CONCLUSIONES Y RECOMENDACIONES TÉCNICAS

| Ámbito | Hallazgo Acreditado | Recomendación Arquitectónica |
| :--- | :--- | :--- |
| **`REPARTIR COSTE`** | Validaciones `NO,SI` en las 3 residencias. Celdas vacías en XLSM. | Campo nullable `dbText(2)` en Access con regla de validación de dominio. No requiere backfill. |
| **`TAG` y `R`** | Formato `General`, sin fórmulas, sin decimales ni texto alfanumérico. | Campo nullable `dbInteger` (DAO Type 3, 2 bytes) para enteros $\le 4$ dígitos. No usar `dbText` ni `dbMemo`. |
| **Facturas Snapshot** | 651 facturas positivas (años 2025-2026). 10 facturas no constan en `Ordenes` ni en `Historico`. | Incorporar las **10 facturas huérfanas** al seed/backfill de `RegistroNumerosFactura` para evitar duplicidad futura. |
| **Colisiones Snapshot** | 13 facturas coinciden en año/número con `Ordenes` pero con distinto titular/orden. | Mantener `Ordenes` como verdad activa vigente; no sobreescribir expedientes activos con datos del snapshot. |
| **Cola `SYS_COLA_SYNC`** | 0 transacciones de datos. PENDIENTE = 0, ERROR = 0. | Transición limpia sin migración de cola en el equipo activo. |
| **WAL Local** | Inexistente/purgado en `SALÓN-PC`. Otras estaciones no accesibles. | Transición autorizada en local; verificar conexión en otros puestos antes de la sustitución del libro. |

---

## 6. CONFIRMACIÓN FINAL DE INTEGRIDAD (CERO MODIFICACIONES)

- **Declaración:** Durante la ejecución de la presente orden, el inspector técnico **NO ha modificado, editado, guardado ni alterado** ningún archivo `.xlsm`, base de datos `.accdb`, módulo VBA, tabla, registro, índice, hoja técnica ni archivo WAL.
- **Verificación de Git:** Comprobado mediante `git status --short`; el estado del repositorio es estrictamente idéntico al momento inicial previo al inicio de la orden.
- **Control de Puerta:** El presente informe y las fuentes analizadas cumplen al 100% el protocolo de codificación, ausencia de mojibake y la regla canónica literal `GIJÓN`.
