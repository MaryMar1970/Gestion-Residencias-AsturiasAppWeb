# RESIDENCIAAPP — INFORME COMPLEMENTARIO DE EVIDENCIAS (ORDEN 1Q-A)
## Comprobaciones Complementarias para el Diseño Arquitectónico

- **Fecha de Emisión**: 2026-11-17
- **Carácter de la Intervención**: Estrictamente lectura forense, comprobación de integridad y verificación técnica de solo lectura.
- **Destinatario**: Dirección Arquitectónica de ResidenciaApp (ChatGPT).
- **Estado de las Restricciones Absolutas**: Cumplimiento del 100% (cero modificaciones de código, cero pseudocódigo, cero propuestas de diseño, regla literal GIJÓN respetada, RGPD protegido sin datos personales).

---

## 1. ESQUEMA, ÍNDICES Y REGISTROS DE LA TABLA `ContadorFacturas` (SECCIÓN 3)

Mediante inspección DAO de solo lectura sobre `H:\ResidenciaBD\Residencia_BE.accdb`, se han verificado los campos, tipos, propiedades, índices y registros actuales de la tabla `ContadorFacturas`:

### 1.1. Campos y Propiedades Físicas

| Pos | Nombre Campo | Tipo DAO | Código | Tamaño | Requerido | AllowZeroLen | Autonumérico | Atributos DAO |
| :---: | :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| 0 | `Id` | `dbLong` | 4 | 4 | False | False | True | 17 |
| 1 | `Residencia` | `dbText` | 10 | 20 | True | False | False | 2 |
| 2 | `Ejercicio` | `dbLong` | 4 | 4 | True | False | False | 1 |
| 3 | `UltimoNumero` | `dbLong` | 4 | 4 | True | False | False | 1 |

### 1.2. Índices y Claves

| Nombre del Índice | Campos Integrantes | Unique | Primary | Required | IgnoreNulls |
| :--- | :--- | :---: | :---: | :---: | :---: |
| `idx_CF_ResEjer` | `Residencia, Ejercicio` | True | False | False | False |
| `Index_28B8F692_E7C9_49A9` | `Id` | True | True | True | False |

### 1.3. Comprobaciones de Unicidad y Registros Actuales

1. **Existencia de Unicidad sobre `(Residencia, Ejercicio)`**: **SÍ existe** mediante el índice `idx_CF_ResEjer` (`Unique = True`). La base de datos impide a nivel de motor registrar dos filas con la misma combinación de sede y ejercicio.
2. **Registros Duplicados**: La consulta `SELECT Residencia, Ejercicio, COUNT(*) FROM ContadorFacturas GROUP BY Residencia, Ejercicio HAVING COUNT(*) > 1` devuelve **0 duplicados**.
3. **Registros Existentes en la Base de Datos**:

| Id | Residencia | Ejercicio | UltimoNumero |
| :---: | :--- | :---: | :---: |
| 1 | `GIJON` | 2026 | 192 |
| 3 | `OVIEDO` | 2026 | 1032 |
| 2 | `SOTO` | 2026 | 20 |

4. **Identificación de Registros Faltantes por Sede y Ejercicio**:
   - **Hecho acreditado**: Solo existen registros para el ejercicio **2026** (GIJON: 192, OVIEDO: 1032, SOTO: 20).
   - **Comportamiento ante otros ejercicios**: No existen registros inicializados para ningún otro ejercicio (p. ej. 2024, 2025, 2027). En el código de `modDatabase.ObtenerSiguienteNumFactura` (líneas 2728-2744), la actualización se ejecuta mediante `UPDATE ContadorFacturas SET UltimoNumero = UltimoNumero + 1 WHERE Residencia = '...' AND Ejercicio = ...`. Si no existe un registro previo para ese ejercicio, la sentencia `UPDATE` afecta a 0 filas y no crea ningún registro, provocando que la asignación falle tras 5 reintentos con el mensaje `"Error al obtener número de factura tras varios intentos"`.

---

## 2. INTEGRIDAD Y COLISIONES DEL CAMPO `Id` (SECCIÓN 4)

Se ha analizado la integridad técnica del campo `Id` en `Ordenes` y `Ordenes_Historico` mediante consultas DAO de solo lectura:

### 2.1. Métricas de Integridad por Tabla

| Métrica Analizada | Tabla `Ordenes` (Activa) | Tabla `Ordenes_Historico` | Método de Obtención DAO |
| :--- | :---: | :---: | :--- |
| **Número Total de Registros** | 2,127 | 17 | `SELECT COUNT(*) FROM <Tabla>` |
| **Valores `Id` Nulos** | 0 | 0 | `SELECT COUNT(*) FROM <Tabla> WHERE Id IS NULL` |
| **Valores `Id` Distintos** | 2,127 | 17 | `SELECT COUNT(*) FROM (SELECT DISTINCT Id FROM <Tabla>)` |
| **Valores `Id` Repetidos** | 0 | 0 | `SELECT Id, COUNT(*) FROM <Tabla> GROUP BY Id HAVING COUNT(*) > 1` |
| **Rango de Valores `Id`** | Min: 2189 / Max: 4346 | Min: 4308 / Max: 4342 | `SELECT MIN(Id), MAX(Id) FROM <Tabla>` |
| **Duplicados Internos** | **0 (Ninguno)** | **0 (Ninguno)** | Verificación sobre índice y agrupación DAO |

### 2.2. Coincidencia de Valores `Id` entre `Ordenes` y `Ordenes_Historico`

- **Consulta ejecutada**: `SELECT O.Id FROM Ordenes O INNER JOIN Ordenes_Historico H ON O.Id = H.Id`
- **Número de valores `Id` coincidentes entre ambas tablas**: **0 (Cero)**.
- **Explicación técnica del resultado**: Aunque el rango numérico de `Id` en `Ordenes_Historico` (4.308 a 4.342) queda contenido dentro del rango global de `Ordenes` (2.189 a 4.346), no existe ningún registro activo que comparta el mismo valor de `Id` con un registro histórico. La razón es que el procedimiento de archivado (`EliminarOrdenBD` / `EliminarOrdenesLoteBD`) copia el registro mediante `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes` (preservando el valor numérico de `Id`), e inmediatamente después ejecuta `DELETE FROM Ordenes WHERE NumOrden = ...` con ese mismo `Id`. Por tanto, el identificador técnico `Id` **distingue unívocamente a cada registro en el conjunto global de ambas tablas**. Las 9 colisiones demostradas en la Orden 1P ocurren sobre la clave funcional `(Residencia, NumOrden)`, donde dos huéspedes distintos tienen números de orden idénticos pero valores de `Id` técnicos diferentes.

---

## 3. COMPROBACIÓN DE LAS FECHAS DE REFERENCIA (SECCIÓN 5)

### 3.1. Correspondencia Literal y Bidireccional de Fechas

1. **Columna B $\leftrightarrow$ `FechaPeticion`**:
   - En `SincronizarFilaHaciaAccess` (`modDatabase.bas`, línea 1898):  
     `AsignarCampoFecha rs, "FechaPeticion", rowVals(1, 2)` (columna 2 = Col B).
   - En `SincronizarHojaDesdeAccess` (`modDatabase.bas`):
     - GIJÓN (línea 3473): `arrGijon1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")`  
     - SOTO (línea 3413): `arrSoto1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")`  
     - OVIEDO (línea 3352): `arrOviedo1(r, 2) = ObtenerValorCampo(rs, "FechaPeticion")`  
   - En `ActualizarOrden` (`modDatabase.bas`, líneas 2554-2557):  
     `If LCase$(campo) = ... Or LCase$(campo) = "fechapeticion" Then setPart = "[" & campo & "] = " & FormatearFechaSQL(CDate(valor))`.
2. **Columna M $\leftrightarrow$ `FechaSalida`**:
   - En `SincronizarFilaHaciaAccess` (`modDatabase.bas`, línea 1919):  
     `AsignarCampoFecha rs, "FechaSalida", rowVals(1, 13)` (columna 13 = Col M).
   - En `SincronizarHojaDesdeAccess` (`modDatabase.bas`):
     - GIJÓN (línea 3488): `arrGijon1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")`  
     - SOTO (línea 3428): `arrSoto1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")`  
     - OVIEDO (línea 3367): `arrOviedo1(r, 13) = ObtenerValorCampo(rs, "FechaSalida")`  
   - En `ActualizarOrden` (`modDatabase.bas`, líneas 2554-2557):  
     `If LCase$(campo) = ... Or LCase$(campo) = "fechasalida" Then setPart = "[" & campo & "] = " & FormatearFechaSQL(CDate(valor))`.
3. **Identidad entre Procedimientos**:
   - Se acredita que `SincronizarFilaHaciaAccess` (Excel $\rightarrow$ Access) y `SincronizarHojaDesdeAccess` (Access $\rightarrow$ Excel) emplean exactamente los mismos campos para las columnas B y M.
4. **Ausencia de Conversiones o Sustituciones que Alteren el Año**:
   - `AsignarCampoFecha` (`modDatabase.bas`, líneas 994-1002) ejecuta `rs(nombreCampo).Value = CDate(valor)` si `IsDate(valor)` es verdadero, o `Null` si está vacío.
   - `FormatearFechaSQL` (`modDatabase.bas`, línea 698) formatea la fecha mediante `#yyyy-mm-dd#` sin modificar año, mes ni día.
   - En Access, ambos campos son de tipo `dbDate` (8 bytes), almacenando fechas canónicas sin manipulación de dígitos.

### 3.2. Catálogo Completo de Usos de `Year(Now)` y `Year(Date)` en el Proyecto

Se han localizado 9 apariciones en todo el código fuente del proyecto, clasificadas según su función:

| Archivo | Línea | Procedimiento | Código Literal | Función / Contexto de Uso |
| :--- | :---: | :--- | :--- | :--- |
| `AsignarNumFactura.bas` | 188 | `AsignarFacturaPorColumna` | `modDatabase.ObtenerSiguienteNumFactura(resCode, Year(Now))` | **Asignación automática de factura**: Solicita el siguiente contador pasando el año del reloj del sistema. |
| `AsignarNumFactura.bas` | 640 | `ProcesarCambioManualFactura` | `modDatabase.SincronizarContadorFacturaManual resCode, numFacturaLong, Year(Now)` | **Asignación manual de factura**: Actualiza `ContadorFacturas` pasando el año del reloj del sistema. |
| `modDatabase.bas` | 2789 | `SincronizarContadorFacturaManual` | `If ejercicio <= 0 Then ejercicio = Year(Now)` | **Fallback de sincronización de factura**: Si el parámetro `ejercicio` no se proporciona o es $\le 0$, toma el año del reloj del sistema. |
| `ModuloSeguridadLocal.bas` | 1134 | `ConsolidarFacturaDiferidaFIFO` | `anioEjercicio = Year(Now)` | **Fallback en diferido offline**: Si la fecha local de la transacción offline no es válida, toma el año del reloj del sistema. |
| `FechasPeticion.bas` | 122 | `FormatearFechaPeticion` | `fechaParte = fechaParte & "/" & Year(Date)` | **Entrada rápida de fecha**: Si el usuario introduce día y mes sin año (p. ej. `15/05`), añade el año actual de la máquina. |
| `FechasPeticion.bas` | 147 | `NormalizarFechaPeticion` | `If Len(fechaCruda) < 6 Then fechaCruda = fechaCruda & "/" & Year(Date)` | **Normalización de fecha**: Añade el año actual de la máquina a entradas cortas. |
| `ModuloCalendarioGijon.bas` | 1092 | `ValidarRangoCalendarioGijon` | `If Year(inicioRango) <> Year(Date) Then` | **Validación de cuadrante**: Comprueba que el calendario visual corresponda al año en curso del sistema. |
| `ModuloCalendarioOviedo.bas` | 1406 | `ValidarRangoCalendarioOviedo` | `If Year(inicioRango) <> Year(Date) Then` | **Validación de cuadrante**: Comprueba que el calendario visual corresponda al año en curso del sistema. |
| `ModuloCalendarioSoto.bas` | 1196 | `ValidarRangoCalendarioSoto` | `If Year(inicioRango) <> Year(Date) Then` | **Validación de cuadrante**: Comprueba que el calendario visual corresponda al año en curso del sistema. |

---

## 4. COMPROBACIÓN DEL CONTADOR DE FACTURAS (SECCIÓN 6)

A partir del análisis de los procedimientos vinculados a la facturación, se constata:

1. **Fecha utilizada actualmente para determinar el ejercicio**:
   - En asignación automática (`AsignarFacturaPorColumna`, línea 188): Se utiliza exclusivamente **`Year(Now)`**.
   - En modificación manual (`ProcesarCambioManualFactura`, línea 640): Se utiliza exclusivamente **`Year(Now)`**.
   - En conciliación offline (`ConsolidarFacturaDiferidaFIFO`, líneas 1131-1135): Se utiliza **`Year(CDate(fechaHoraLocal))`** (marca temporal del reloj del puesto en el instante en que se encoló la operación offline) con fallback a **`Year(Now)`**.
   - **Conclusión**: En ningún procedimiento se utiliza `Year(FechaSalida)` (columna M).
2. **Consulta del Máximo de `NumFactura`**:
   - En `modDatabase.ObtenerSiguienteNumFactura` (línea 2711):  
     `SELECT MAX(NumFactura) FROM Ordenes WHERE Residencia = '...'`  
     **Hecho acreditado**: No se filtra por ningún ejercicio ni fecha. Se extrae el máximo absoluto de la tabla `Ordenes` para esa sede.
3. **Tablas Consultadas**:
   - `ObtenerSiguienteNumFactura` (línea 2711) consulta **únicamente `Ordenes`**. No consulta `Ordenes_Historico`.
   - `ExisteFacturaEnBD` (línea 2901) consulta **únicamente `Ordenes`** (`SELECT COUNT(*) FROM Ordenes WHERE Residencia = '...' AND NumFactura = ...`). No consulta `Ordenes_Historico`.
4. **Criterio de Ejercicio entre Factura Manual y Automática**:
   - Ambas utilizan exactamente el mismo criterio: invocan a `modDatabase` pasando el valor devuelto por `Year(Now)`.
5. **Comportamiento de la Factura Diferida Offline**:
   - `ConsolidarFacturaDiferidaFIFO` **no lee la fecha de salida** de la orden. Evalúa la cadena `fechaHoraLocal` registrada en la cola WAL en el momento de la pre-reserva:
     ```vba
     If IsDate(fechaHoraLocal) Then
         anioEjercicio = Year(CDate(fechaHoraLocal))
     Else
         anioEjercicio = Year(Now)
     End If
     ```
6. **Recepción y Uso del Ejercicio en Modificaciones de `ContadorFacturas`**:
   - Los dos procedimientos que modifican la tabla (`ObtenerSiguienteNumFactura` y `SincronizarContadorFacturaManual`) reciben el parámetro `ejercicio As Integer` (o Long) y lo aplican en la cláusula `WHERE Ejercicio = ...`.
   - Sin embargo, los llamadores les suministran siempre el año del reloj (`Year(Now)` o `Year(fechaHoraLocal)`), y como se acreditó en la Sección 1, si la tabla no contiene una fila previa para ese ejercicio, la sentencia `UPDATE` afecta a 0 registros y la operación falla.

---

## 5. CLASIFICACIÓN DE COLUMNAS DE DATOS Y FÓRMULAS POR RESIDENCIA (SECCIÓN 7)

A partir de la inspección física de las hojas XML de `BDAS_Multiusuario.xlsm` (`sheet2.xml`, `sheet23.xml` y `sheet34.xml`), se documenta la clasificación de columnas en cada residencia para permitir el diseño arquitectónico del recuento de 15 columnas sin interferencia de fórmulas:

### 5.1. Mapa Integral de Columnas por Residencia

| Col | Letra | Encabezado en GIJÓN | Encabezado en SOTO | Encabezado en OVIEDO | Naturaleza en GIJÓN | Naturaleza en SOTO | Naturaleza en OVIEDO |
| :---: | :---: | :--- | :--- | :--- | :---: | :---: | :---: |
| 1 | **A** | Nº ORDEN | Nº ORDEN | Nº ORDEN | **DATO (Req)** | **DATO (Req)** | **DATO (Req)** |
| 2 | **B** | FECHA PETICION | FECHA PETICION | FECHA PETICION | **DATO** | **DATO** | **DATO** |
| 3 | **C** | NÚM FACT. | NÚM FACT. | NÚM FACT | **DATO** | **DATO** | **DATO** |
| 4 | **D** | FINALIDAD | FINALIDAD | FINALIDAD | **DATO** | **DATO** | **DATO** |
| 5 | **E** | EMPLEO | EMPLEO | EMPLEO | **DATO** | **DATO** | **DATO** |
| 6 | **F** | SITUACION | SITUACION | SITUACION | **DATO** | **DATO** | **DATO** |
| 7 | **G** | EVALUACIÓN | EVALUACIÓN | EVALUACIÓN | **DATO** | **DATO** | **DATO** |
| 8 | **H** | COMISIÓN | TURNO | COMISIÓN | **DATO** | **DATO** | **DATO** |
| 9 | **I** | DNI | DNI | DNI | **DATO** | **DATO** | **DATO** |
| 10 | **J** | RANGO | RANGO | RANGO | **DATO** | **DATO** | **DATO** |
| 11 | **K** | Nombre | Nombre | Nombre | **DATO** | **DATO** | **DATO** |
| 12 | **L** | ENTRADA | ENTRADA | ENTRADA | **DATO** | **DATO** | **DATO** |
| 13 | **M** | SALIDA | SALIDA | SALIDA | **DATO** | **DATO** | **DATO** |
| 14 | **N** | DIAS USO | DIAS USO | DIAS USO | **DATO** | **DATO** | **DATO** |
| 15 | **O** | PAX | PAX | PAX | **DATO** | **DATO** | **DATO** |
| 16 | **P** | RESOLUCION | RESOLUCION | RESOLUCION | **DATO** | **DATO** | **DATO** |
| 17 | **Q** | HAB. IND. | PRECIO APTO | HAB. IND. | **DATO** | *FÓRMULA* | **DATO** |
| 18 | **R** | HAB. DOBLE | SUPLE OCUPAN | HAB. DOBLE | **DATO** | *FÓRMULA* | **DATO** |
| 19 | **S** | NÚM HAB. | APARTAMENTO | CAMA SUPLE. | **DATO** | **DATO** | **DATO** |
| 20 | **T** | DTO. FAM. NUM. | DTO. FAM. NUM. | NÚM HAB. | **DATO** | **DATO** | **DATO** |
| 21 | **U** | IMPORTE | IMPORTE | DTO. FAM. NUM. | *FÓRMULA* | *FÓRMULA* | **DATO** |
| 22 | **V** | TELEFONO | TELEFONO | IMPORTE | **DATO** | **DATO** | *FÓRMULA* |
| 23 | **W** | DIRECCIÓN | DIRECCIÓN | TELEFONO | **DATO** | **DATO** | **DATO** |
| 24 | **X** | CP | CP | DIRECCIÓN | **DATO** | **DATO** | **DATO** |
| 25 | **Y** | POBLACIÓN | POBLACIÓN | CP | **DATO** | **DATO** | **DATO** |
| 26 | **Z** | PROVINCIA | PROVINCIA | POBLACIÓN | **DATO** | **DATO** | **DATO** |
| 27 | **AA** | PAGADO | PAGADO | PROVINCIA | **DATO** | **DATO** | **DATO** |
| 28 | **AB** | GRABACIÓN SOLICITUD | GRABACIÓN SOLICITUD | PAGADO | *FÓRMULA* | *FÓRMULA* | **DATO** |
| 29 | **AC** | COPIAR TEXTO EMAIL | COPIAR TEXTO EMAIL | REPARTIR COSTE HAB. | *FÓRMULA* | *FÓRMULA* | **DATO** |
| 30 | **AD** | REPARTIR COSTE HAB. | REPARTIR COSTE APTO. | TAG | **DATO** | **DATO** | **DATO** |
| 31 | **AE** | *(Sin uso)* | *(Sin uso)* | R | - | - | **DATO** |
| 32 | **AF** | *(Sin uso)* | *(Sin uso)* | GRABACIÓN SOLICITUD | - | - | *FÓRMULA* |
| 33 | **AG** | *(Sin uso)* | *(Sin uso)* | COPIAR TEXTO EMAIL | - | - | *FÓRMULA* |

### 5.2. Resumen Numérico de Columnas de Datos frente a Fórmulas

| Concepto | RESIDENCIA GIJÓN | RESIDENCIA SOTO | RESIDENCIA OVIEDO |
| :--- | :---: | :---: | :---: |
| **Columnas de Datos Introducidos por Usuario** | **26 columnas**<br>(A:T, V:AA, AD) | **24 columnas**<br>(A:P, S:T, V:AA, AD) | **29 columnas**<br>(A:U, W:AE) |
| **Columnas con Fórmulas Calculadas** | **3 columnas**<br>(U, AB, AC) | **5 columnas**<br>(Q, R, U, AB, AC) | **3 columnas**<br>(V, AF, AG) |
| **Total de Columnas Físicas Operativas** | 29 columnas (A a AD) | 29 columnas (A a AD) | 33 columnas (A a AG) |

### 5.3. Columnas que Pueden Estar Vacías con Datos Válidos en Otras

1. **Columna C (`NÚM FACT.`)**: Está vacía en todas las solicitudes pendientes de pago, denegadas o canceladas. Solo se cumplimenta tras efectuar el cobro.
2. **Columnas Q y R (`HAB. IND.` y `HAB. DOBLE` en GIJÓN y OVIEDO)**: Se cumplimentan de forma mutuamente excluyente según la tipología de alojamiento solicitada (si solicita individual, doble queda vacía; si solicita doble, individual queda vacía).
3. **Columna S (`CAMA SUPLE.` en OVIEDO)**: Solo se cumplimenta si el huésped solicita cama supletoria; permanece vacía en la mayoría de reservas.
4. **Columna T en GIJÓN/SOTO y Columna U en OVIEDO (`DTO. FAM. NUM.`)**: Solo se cumplimenta si el solicitante acredita condición de familia numerosa (`DTO 10%` o `DTO 20%`); en caso contrario está vacía.
5. **Columnas Postales y de Contacto (Teléfono, Dirección, CP, Población, Provincia)**: En peticiones iniciales o reservas telefónicas rápidas, el operador frecuentemente cumplimenta únicamente el DNI, Nombre y Teléfono, dejando Dirección, CP y Población vacíos.
6. **Columna de Pago (`AA` en GIJÓN/SOTO, `AB` en OVIEDO)**: Permanece vacía hasta la confirmación de la pasarela de pago o recepción de la transferencia.
7. **Columna de Reparto de Coste (`AD` en GIJÓN/SOTO, `AC` en OVIEDO)**: Solo se cumplimenta con `SI` cuando intervienen múltiples pagadores; en el resto permanece vacía.

### 5.4. Diferencias Físicas Específicas entre Sedes

1. **Columna H**: En GIJÓN y OVIEDO almacena `COMISIÓN` (`SI` / `NO`); en SOTO almacena `TURNO` (`T-1`, `T-2`, etc.).
2. **Columnas Q y R (17 y 18)**: En GIJÓN y OVIEDO son columnas de datos de usuario (`HAB. IND.` y `HAB. DOBLE`); en SOTO son **fórmulas de cálculo de precios** (`PRECIO APTO` y `SUPLE OCUPAN`).
3. **Columna S (19)**: En GIJÓN almacena `NÚM HAB.`; en SOTO almacena `APARTAMENTO`; en OVIEDO almacena `CAMA SUPLE.`.
4. **Columna T (20)**: En GIJÓN y SOTO almacena `DTO. FAM. NUM.`; en OVIEDO almacena `NÚM HAB.`.
5. **Columna U (21)**: En GIJÓN y SOTO es la **fórmula de `IMPORTE`**; en OVIEDO es la columna de datos `DTO. FAM. NUM.`.
6. **Columna V (22)**: En GIJÓN y SOTO almacena `TELEFONO`; en OVIEDO es la **fórmula de `IMPORTE`**.
7. **Desplazamiento general en OVIEDO**: Debido a la existencia de la columna `CAMA SUPLE.` (S), todos los datos postales de OVIEDO se encuentran desplazados una columna a la derecha respecto a GIJÓN y SOTO (Teléfono en W, Dirección en X, CP en Y, Población en Z, Provincia en AA, Pagado en AB).
8. **Fórmulas de generación de PDF y Email**: En GIJÓN y SOTO residen en las columnas AB (28) y AC (29); en OVIEDO residen en las columnas AF (32) y AG (33).

---

## 6. CONCLUSIONES OBJETIVAS Y REFERENCIAS A LAS FUENTES (SECCIÓN 8)

1. **Unicidad del identificador técnico `Id`**: Se confirma documentalmente que no existe colisión de `Id` entre `Ordenes` y `Ordenes_Historico` (0 coincidencias). Cada registro tiene un `Id` único en el universo completo de la base de datos. Las colisiones entre expedientes obedecen exclusivamente a la clave funcional `(Residencia, NumOrden)`.
2. **Unicidad e índices en `ContadorFacturas`**: La tabla posee índice único sobre `(Residencia, Ejercicio)` (`idx_CF_ResEjer`). No existen registros duplicados. Sin embargo, solo existen contadores para el ejercicio 2026; la consulta para cualquier otro ejercicio no crea el registro y falla.
3. **Invariabilidad de las fechas de referencia B y M**: La correspondencia de Columna B con `FechaPeticion` y Columna M con `FechaSalida` es 100% simétrica y directa en todos los módulos de sincronización. Las fechas se conservan sin alteraciones de año en el almacenamiento `dbDate`.
4. **Uso de `Year(Now)` en facturación**: Todo el subsistema de facturación actual (automática, manual y diferida) depende de `Year(Now)` o del reloj de la máquina, omitiendo completamente `FechaSalida` (Col M) y consultando `MAX(NumFactura)` en `Ordenes` sin filtro de ejercicio.
5. **Diferencias físicas de columnas**: Se aporta el catálogo detallado de columnas de datos frente a fórmulas para GIJÓN (26 datos / 3 fórmulas), SOTO (24 datos / 5 fórmulas) y OVIEDO (29 datos / 3 fórmulas), acreditando los desplazamientos de columnas y las celdas opcionales para el futuro diseño del recuento de 15 columnas.

---
**Fin del Informe Complementario — Orden 1Q-A.**
