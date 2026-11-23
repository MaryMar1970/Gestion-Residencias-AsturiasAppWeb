# RESIDENCIAAPP — INFORME DE EVIDENCIAS TÉCNICAS (ORDEN 1T)
## EXTRACCIÓN DE EVIDENCIAS PARA CERRAR LA IDENTIDAD DEL EXPEDIENTE, LA PERSISTENCIA OFFLINE Y LA FACTURACIÓN

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Diseño arquitectónico de numeración, facturación y archivado  
**Naturaleza:** Inspección estática y consultas DAO/OLEDB estrictamente de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha y hora de emisión:** 2026-11-19 19:05:00 UTC+1  
**Documento de referencia previa:** Orden 1S ([RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md](file:///C:/Users/jaime_ntymfau/.gemini/antigravity/brain/1bda4f34-55a5-4d42-b9bf-6092244693f2/RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md))

---

## 1. CONTROL DE INTEGRIDAD DE LAS FUENTES Y VERIFICACIÓN FORENSE

Se ha procedido al cálculo y verificación de los hashes criptográficos SHA-256 de las fuentes autorizadas antes de iniciar cualquier análisis técnico:

| Componente | Ruta física de inspección | Tamaño (bytes) | Fecha y hora de modificación | Hash SHA-256 verificado | Coincidencia con Orden 1T |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Libro Excel Activo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 1.157.075 bytes | 2026-11-16 23:44:03 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` | **100% EXACTA** |
| **Módulo VBA de Base de Datos** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 134.120 bytes (3.090 líneas) | 2026-11-16 23:44:03 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | **100% EXACTA** |
| **Backend Access de Pruebas** | `H:\ResidenciaBD\Residencia_BE.accdb` | 2.502.656 bytes | 2026-11-17 19:40:27 | `9617f99795b8266c2abb91b479bf64b9ece449ad15885ce40edbe62639dbbee9` | **100% EXACTA** |

Todas las operaciones se han realizado en modo estrictamente de solo lectura, sin alteración de metadatos, esquemas ni registros de datos.

---

## 2. COMPROBACIÓN A: MAPA FACTUAL DEL CICLO DE VIDA DE LA IDENTIDAD TÉCNICA Y DE NEGOCIO

### 2.1. Inexistencia formal de `CrearOrden`
Se ha realizado un rastreo exhaustivo en la totalidad de módulos estándar (`.bas`), módulos de clase (`.cls`), hojas operativas y formularios del proyecto:
- **Resultado:** El procedimiento `CrearOrden` **NO EXISTE** en el código base (0 coincidencias en 100% del proyecto). Dicha denominación figuraba únicamente en esquemas y propuestas documentales teóricas previas.

### 2.2. Procedimientos reales de creación e inserción de expedientes
La creación e inserción física de registros en la tabla `Ordenes` de Access se articula en producción a través de los siguientes procedimientos localizados en `modDatabase.bas`:

1. **`InsertarOrdenAtomica` (`modDatabase.bas`, líneas 846-950):**
   - **Mecanismo:** Inicia una transacción ADODB explícita (`cn.BeginTrans`).
   - **Asignación de número:** Ejecuta `SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '...'` y calcula `nuevoNumOrden = maxActual + 1`.
   - **Inserción física:** Ejecuta `INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, ...) VALUES (...)`.
   - **Tratamiento del identificador técnico:** **Omite por completo el campo `Id`**. Access genera automáticamente el autonumérico `Counter`, pero la función no lo recupera ni lo retorna. La función devuelve exclusivamente `nuevoNumOrden As Long`.

2. **`IntentarReservarOrdenBD` (`modDatabase.bas`, líneas 1281-1375):**
   - **Mecanismo:** Diseñado para la entrada manual directa de un `NumOrden` específico en la columna A.
   - **Control de concurrencia:** Adquiere un Lease Lock sobre la clave `resCode & ":" & numOrdenPropuesto`.
   - **Inserción física:** Ejecuta un `INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion...)` insertando un registro preliminar.
   - **Tratamiento del identificador técnico:** No genera ni recupera ningún `Id` de instancia.

3. **`ReservarSiguienteOrdenLibreBD` (`modDatabase.bas`, líneas 1403-1470):**
   - **Mecanismo:** Detecta el primer hueco numérico libre o recurre a `MAX(NumOrden) + 1` en Access.
   - **Inserción física:** Inserta el registro preliminar en `Ordenes` e invoca `InsertarAuditoriaSegura`.
   - **Tratamiento del identificador técnico:** Retorna exclusivamente el `NumOrden As Long`.

4. **`AsegurarOrdenExiste` (`modDatabase.bas`, líneas 1698-1732):**
   - **Mecanismo:** Invocado por `SincronizarFilaHaciaAccess`. Ejecuta `SELECT COUNT(*) FROM Ordenes WHERE NumOrden = ... AND Residencia = ...`. Si el conteo es 0, ejecuta un `INSERT` con el `NumOrden` y la `Residencia` para asegurar la existencia del registro antes de actualizar sus campos.

5. **`SincronizarFilaHaciaAccess` (`modDatabase.bas`, líneas 1737-1995):**
   - **Mecanismo:** Lee las columnas 1 a 28 de la fila de Excel en un array de memoria. Adquiere el Lease Lock y realiza un `Recordset.Open` con `LockType = adLockOptimistic`.
   - **Si el registro no existe y `permitirAlta = True`:** Ejecuta `rs.AddNew`, asignando `rs("NumOrden") = numOrden` y el resto de campos.
   - **Tratamiento del identificador técnico:** No lee `rs("Id")` ni lo traslada a la hoja Excel.

6. **`SincronizarCambioTargetHaciaAccess` (`modDatabase.bas`, líneas 2307-2445):**
   - **Mecanismo:** Disparado celda a celda por `Worksheet_Change`. Traduce la columna a campo Access mediante `ObtenerCampoBDDesdeColumna` y ejecuta un `UPDATE Ordenes SET [Campo] = ... WHERE NumOrden = ... AND Residencia = ...`.

### 2.3. Resumen factual de la identidad
- **Identidad en el ciclo online:** La única clave de enlace entre Excel y Access es el par de negocio `(Residencia, NumOrden)`.
- **Identidad técnica en Access (`Id`):** El campo autonumérico `Id` de Access es puramente interno del motor Jet/ACE; **jamás se proyecta hacia Excel, no reside en ninguna columna de las hojas ni se conserva en la memoria local**.
- **Identidad de instancia (`IdInstanciaOrden`):** **No existe implementación alguna** de un identificador único permanente de instancia en el código vigente.
- **Consecuencia en el archivado y borrado:** Como se constató en las Órdenes 1O y 1P, al no existir `IdInstanciaOrden` ni filtro por año en las sentencias de archivado (`DELETE FROM Ordenes_Historico WHERE Residencia = '...' AND NumOrden = ...`), la reutilización de un `NumOrden` destruye físicamente cualquier expediente anterior que hubiera compartido dicho número.

---

## 3. COMPROBACIÓN B: RECORRIDO COMPLETO DE LA COLA OFFLINE (`SYS_COLA_SYNC`) Y DEL DIARIO WAL

### 3.1. Estructura exacta de la cola `SYS_COLA_SYNC`
La persistencia de contingencia local se estructura en la hoja oculta `SYS_COLA_SYNC`, gobernada por la enumeración `eColCola` en `ModuloSeguridadLocal.bas` (líneas 39-52):

| Columna | Nombre lógico del Enum | Tipo de dato / Formato | Propósito en el sistema |
| :---: | :--- | :--- | :--- |
| **A (1)** | `colIDTransaccion` | `Long` (Secuencial 1, 2, 3...) | Orden estricto FIFO de la operación. |
| **B (2)** | `colFechaHoraLocal` | `String` (`yyyy-mm-dd hh:nn:ss`) | Marca temporal local en el momento de encolar. |
| **C (3)** | `colUsuario` | `String` (Nombre de usuario Windows) | Identidad del operador local. |
| **D (4)** | `colResidencia` | `String` (`GIJON`, `SOTO`, `OVIEDO`) | Sede de la orden afectada. |
| **E (5)** | `colNumOrden` | `Long` (Definitivo o provisional `900xxx`) | Clave de negocio del expediente. |
| **F (6)** | `colTipoOperacion` | `String` | Tipos válidos: `UPDATE_CAMPO`, `ASIGNAR_FACTURA_DIFERIDA`, `PRERESERVA`, `ELIMINAR_ORDEN`, `LOG_OFFLINE`. |
| **G (7)** | `colCampoModificado` | `String` | Nombre físico de la columna en Access (`EstadoPago`, `HabitacionesAsignadas`, etc.). |
| **H (8)** | `colValorAnterior` | `String` (Pre-Image) | Valor previo a la edición para trazabilidad / rollback. |
| **I (9)** | `colValorNuevo` | `String` (Post-Image) | Valor final introducido por el usuario. |
| **J (10)** | `colEstadoSync` | `String` | Estados: `PENDIENTE`, `SINCRONIZANDO`, `CONCILIADO`, `CONFLICTO_HAB`, `ERROR`. |
| **K (11)** | `colReintentos` | `Integer` | Contador de fallos durante la sincronización. |
| **L (12)** | `colUUID` | `String` (Token pseudoaleatorio) | Token de idempotencia de la fila de transacción. |
| **M (13)** | `colCRC` | `Long` | CRC16 de integridad sobre los campos de la fila. |

### 3.2. Mecanismo de persistencia en disco: Diario WAL
- **Ubicación física:** `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\SYS_COLA_SYNC_<COMPUTERNAME>.wal`
- **Serialización:** Formato de texto plano, una línea por transacción encolada, delimitada por barras verticales (`|`).
- **Saneamiento de delimitadores:** Función `SanitizarPipe` (`ModuloSeguridadLocal.bas`, líneas 1643-1651):
  - Sustituye `|` por `{PIPE}`.
  - Sustituye `vbCrLf` por `{CRLF}`.
  - Sustituye `;` por `{SEMI}`.
- **Integridad:** Función `CalcularCRC16` calculada sobre la cadena canónica:
  `idTrans & "|" & uuidOp & "|" & resNorm & "|" & numOrden & "|" & tipoOperacion & "|" & campoModificado & "|" & valorNuevo`.
- **Recuperación al inicio:** En la apertura del libro (`ThisWorkbook.cls`), el procedimiento `CargarColaDesdeWAL` reconstruye la hoja `SYS_COLA_SYNC` si el archivo WAL local contiene transacciones pendientes no reflejadas.

### 3.3. Coalescencia de operaciones offline
En `ModuloSeguridadLocal.EncolarOperacionOffline` (líneas 342-357):
- Si una operación es de tipo `UPDATE_CAMPO` o `ASIGNAR_FACTURA_DIFERIDA`, se realiza una búsqueda lineal en las filas existentes con `colEstadoSync = "PENDIENTE"`.
- Si coincide `(colResidencia, colNumOrden, colCampoModificado, colTipoOperacion)`, el sistema **no genera una nueva fila en la cola**: actualiza la celda `colValorNuevo` de la fila ya existente y conserva el `colValorAnterior` original. Se recalcula el CRC16 y se añade la línea de actualización al WAL.

### 3.4. Ciclo de las órdenes provisionales (`900xxx`)
1. **Generación offline:** Al crear una orden sin conexión en la columna A, `GestionarEntradaNumOrdenHoja` detecta el modo contingencia y asigna un número provisional `900xxx` (a partir de 900001 escaneando la columna A local).
2. **Encolado:** Se encola una operación `PRERESERVA` con `colNumOrden = 900xxx`. Todas las ediciones posteriores en otras columnas de esa fila se encolan con `colNumOrden = 900xxx`.
3. **Reconciliación online:** Al recuperarse la red, `ReconciliarColaOfflineFIFO` invoca `ResolverNumOrdenProvisionalEnBD` (`ModuloSeguridadLocal.bas`, líneas 1364-1402):
   - Obtiene un número definitivo en Access mediante `SELECT MAX(NumOrden) + 1`.
   - Inserta la orden en Access con el número definitivo.
   - Actualiza la celda de la Columna A en la hoja Excel.
   - **Actualización en cascada de la cola:** Recorre todas las filas pendientes restantes en `SYS_COLA_SYNC` y sustituye el valor `900xxx` por el nuevo número definitivo tanto en la hoja oculta como en el WAL.

### 3.5. Distinción técnica entre `colUUID` y un identificador permanente de entidad
Se ha inspeccionado la función `GenerarUUID` (`ModuloSeguridadLocal.bas`, líneas 1625-1628):
```vba
Private Function GenerarUUID(ByVal res As String, ByVal seq As Long) As String
    GenerarUUID = Left$(res, 3) & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
                  Format$(seq, "0000") & "-" & Right$("0000" & Hex$(Int(Rnd() * 65535)), 4)
End Function
```
- **Conclusión técnica:** `colUUID` es un identificador de **transacción** (un token de idempotencia para la fila de sincronización). Su función exclusiva es evitar la duplicidad al reintentar el volcado FIFO hacia Access. **No identifica al expediente, no se propaga a la tabla `Ordenes`, no se guarda en Excel y desaparece al purgar la cola de sincronización.**

---

## 4. COMPROBACIÓN C: MAPA FACTUAL DE PERSISTENCIA DE COLUMNAS (EXCEL VS ACCESS)

Se ha cruzado la extracción directa de las cabeceras XML y fórmulas de las hojas operativas de `BDAS_Multiusuario.xlsm` con la implementación del procedimiento `ObtenerCampoBDDesdeColumna` (`modDatabase.bas`, líneas 2241-2299) y `SincronizarFilaHaciaAccess` (líneas 1737-1995).

### 4.1. Tabla completa de correspondencia de columnas por residencia

| Col | Letra | Nombre de Cabecera en Hoja Excel | Campo en Access (`Ordenes`) | ¿Tiene Persistencia en Access? | Observaciones técnicas y comportamiento del código |
| :---: | :---: | :--- | :--- | :---: | :--- |
| **1** | **A** | `Nº ORDEN` | `NumOrden` | **SÍ** | Clave de negocio. Persistido mediante `IntentarReservarOrdenBD` / `InsertarOrdenAtomica`. |
| **2** | **B** | `FECHA PETICION` | `FechaPeticion` | **SÍ** | Formato fecha. Base del Ejercicio de la Orden. |
| **3** | **C** | `Nº FACTURA` | `NumFactura` | **SÍ** | Gestionado por `AsignarNumFactura` y `ProcesarCambioManualFactura`. |
| **4** | **D** | `FINALIDAD` | `Finalidad` | **SÍ** | Lista desplegable / Texto. |
| **5** | **E** | `EMPLEO` | `Empleo` | **SÍ** | Lista desplegable / Texto. |
| **6** | **F** | `SITUACION` | `Situacion` | **SÍ** | Lista desplegable / Texto. |
| **7** | **G** | `EVALUACION` | `Evaluacion` | **SÍ** | Calculada o manual según finalidad. |
| **8** | **H** | `TURNO` | `Turno` | **SÍ** | Turno asignado. |
| **9** | **I** | `DNI` | `DNI` | **SÍ** | Validado contra Lista Negra y formato DNI/NIE. |
| **10** | **J** | `TITULAR/BENEFICIARIO` | `TitularBeneficiario` | **SÍ** | Texto. |
| **11** | **K** | `NOMBRE COMPLETO` | `Nombre` | **SÍ** | ProperCase. |
| **12** | **L** | `F. ENTRADA` | `FechaEntrada` | **SÍ** | Formato fecha. |
| **13** | **M** | `F. SALIDA` | `FechaSalida` | **SÍ** | Formato fecha. Base del Ejercicio de Factura. |
| **14** | **N** | `DIAS` | `DiasEstancia` | **SÍ** | Entero / Días calculados. |
| **15** | **O** | `PERSONAS` | `NumeroPersonas` | **SÍ** | Entero. |
| **16** | **P** | `RESOLUCION` | `Resolucion` | **SÍ** | Estado: CONCEDIDA, DENEGADA, RENUNCIA, etc. |
| **17** | **Q** | *GIJÓN / OVIEDO:* `HAB. IND.`<br>*SOTO:* `APTO. 1 HAB.` | `HabitacionesIndividuales` *(GIJÓN/OVIEDO)*<br>`Apto1Hab` *(SOTO)* | **SÍ** *(GIJÓN/OVIEDO)*<br>**NO** *(SOTO - Fórmula)* | En SOTO es una celda con fórmula (`=SI(...)`). `ObtenerCampoBDDesdeColumna` retorna `""`. |
| **18** | **R** | *GIJÓN / OVIEDO:* `HAB. DOBLE`<br>*SOTO:* `APTO. 2 HAB.` | `HabitacionesDobles` *(GIJÓN/OVIEDO)*<br>`Apto2Hab` *(SOTO)* | **SÍ** *(GIJÓN/OVIEDO)*<br>**NO** *(SOTO - Fórmula)* | En SOTO es una celda con fórmula (`=SI(...)`). `ObtenerCampoBDDesdeColumna` retorna `""`. |
| **19** | **S** | *GIJÓN / SOTO:* `Nº HAB.` / `NUMERO`<br>*OVIEDO:* `CAMAS SUP.` | `HabitacionesAsignadas` *(GIJÓN/SOTO)*<br>`CamasSupletorias` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **20** | **T** | *GIJÓN / SOTO:* `CONFIRMAR RESERVA`<br>*OVIEDO:* `HABITACIONES` | `ConfirmarReserva` *(GIJÓN/SOTO)*<br>`HabitacionesAsignadas` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **21** | **U** | *GIJÓN / SOTO:* `IMPORTE`<br>*OVIEDO:* `CONFIRMAR RESERVA` | *GIJÓN/SOTO:* *Fórmula Excel*<br>*OVIEDO:* `ConfirmarReserva` | **NO** *(GIJÓN/SOTO)*<br>**SÍ** *(OVIEDO)* | En GIJÓN y SOTO es fórmula calculada. `ObtenerCampoBDDesdeColumna` retorna `""`. |
| **22** | **V** | *GIJÓN / SOTO:* `TELEFONO`<br>*OVIEDO:* `IMPORTE` | `Telefono` *(GIJÓN/SOTO)*<br>*OVIEDO:* *Fórmula Excel* | **SÍ** *(GIJÓN/SOTO)*<br>**NO** *(OVIEDO)* | En OVIEDO es fórmula calculada. `ObtenerCampoBDDesdeColumna` retorna `""`. |
| **23** | **W** | *GIJÓN / SOTO:* `CORREO`<br>*OVIEDO:* `TELEFONO` | `Correo` *(GIJÓN/SOTO)*<br>`Telefono` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **24** | **X** | *GIJÓN / SOTO:* `CODIGO POSTAL`<br>*OVIEDO:* `CORREO` | `CodigoPostal` *(GIJÓN/SOTO)*<br>`Correo` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **25** | **Y** | *GIJÓN / SOTO:* `POBLACION`<br>*OVIEDO:* `CODIGO POSTAL` | `Poblacion` *(GIJÓN/SOTO)*<br>`CodigoPostal` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **26** | **Z** | *GIJÓN / SOTO:* `PROVINCIA`<br>*OVIEDO:* `POBLACION` | `Provincia` *(GIJÓN/SOTO)*<br>`Poblacion` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **27** | **AA** | *GIJÓN / SOTO:* `PAGADO`<br>*OVIEDO:* `PROVINCIA` | `EstadoPago` *(GIJÓN/SOTO)*<br>`Provincia` *(OVIEDO)* | **SÍ** | Mapeado a campo de base de datos. |
| **28** | **AB** | *GIJÓN / SOTO:* `GRABACIÓN`<br>*OVIEDO:* `PAGADO` | *GIJÓN/SOTO:* *Fórmula Excel*<br>*OVIEDO:* `EstadoPago` | **NO** *(GIJÓN/SOTO)*<br>**SÍ** *(OVIEDO)* | En GIJÓN/SOTO es fórmula concatenada para volcado. En OVIEDO es campo de datos `EstadoPago`. |
| **29** | **AC** | *GIJÓN / SOTO:* `COPIAR TEXTO EMAIL`<br>*OVIEDO:* `REPARTIR COSTE  HAB.` | *GIJÓN/SOTO:* *Fórmula / Acción*<br>*OVIEDO:* **DATO LOCAL DE USUARIO** | **NO** | En GIJÓN/SOTO es fórmula. En OVIEDO son datos de usuario pero `ObtenerCampoBDDesdeColumna` retorna `""`. **NO TIENE CAMPO EN ACCESS**. |
| **30** | **AD** | *GIJÓN:* `REPARTIR COSTE  HAB.`<br>*SOTO:* `REPARTIR COSTE  APTO.`<br>*OVIEDO:* `TAG` | **DATO LOCAL DE USUARIO** | **NO** | **NO TIENEN CAMPO EN ACCESS**. Datos 100% volátiles en Excel. `ObtenerCampoBDDesdeColumna` retorna `""`. |
| **31** | **AE** | *OVIEDO:* `R` | **DATO LOCAL DE USUARIO** | **NO** | Exclusivo de OVIEDO. `ObtenerCampoBDDesdeColumna` retorna `""`. **NO TIENE CAMPO EN ACCESS**. |
| **32** | **AF** | *OVIEDO:* `GRABACIÓN` | *Fórmula Excel* | **NO** | Fórmula concatenada de control en OVIEDO. Sin persistencia en Access. |
| **33** | **AG** | *OVIEDO:* `COPIAR TEXTO EMAIL` | *Fórmula Excel* | **NO** | Fórmula para plantilla de correo en OVIEDO. Sin persistencia en Access. |

### 4.2. Hallazgo crítico sobre las columnas de usuario sin persistencia
- **Columnas afectadas:**
  - `RESIDENCIA GIJÓN`: Columna 30 (`AD`), encabezado `REPARTIR COSTE  HAB.`.
  - `RESIDENCIA SOTO`: Columna 30 (`AD`), encabezado `REPARTIR COSTE  APTO.`.
  - `RESIDENCIA OVIEDO`: Columna 29 (`AC`), `REPARTIR COSTE  HAB.`; Columna 30 (`AD`), `TAG`; Columna 31 (`AE`), `R`.
- **Comportamiento acreditado en código:**
  1. `SincronizarFilaHaciaAccess` ejecuta estrictamente `arrFila = ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 28)).Value2`. **Solo lee las columnas 1 a 28**, ignorando por completo cualquier columna superior a la 28.
  2. `ObtenerCampoBDDesdeColumna` no incluye ningún mapeo para estas columnas y retorna una cadena vacía `""`.
  3. `SincronizarCambioTargetHaciaAccess` sale de inmediato sin ejecutar ninguna acción en Access (`If Len(campoBD) = 0 Then Exit Sub`).
  4. La cola offline `SYS_COLA_SYNC` jamás encola modificaciones de estas columnas.
  5. Ni la tabla `Ordenes` ni la tabla `Ordenes_Historico` disponen de campos para alojar estos datos.
  6. **Consecuencia funcional:** Los datos introducidos en estas columnas residen exclusivamente en las celdas de la hoja Excel activa. Si el libro se cierra sin guardar, si se ejecuta un refresco completo desde Access, si la fila se archiva a `Ordenes_Historico` o si se borra localmente, **la información de estas columnas se pierde irremediablemente**.

---

## 5. COMPROBACIÓN D: CICLO REAL DE FACTURACIÓN, CONTADOR Y ESQUEMA FÍSICO DAO DE `ContadorFacturas`

### 5.1. Esquema físico DAO exacto de `ContadorFacturas`
Mediante inspección física directa con el motor DAO 12.0 (`DAO.DBEngine.120`) sobre el backend de producción `H:\ResidenciaBD\Residencia_BE.accdb`, se han extraído las especificaciones técnicas rigurosas de la tabla:

| Campo | Tipo DAO (`Type`) | Nombre del Tipo DAO | Tamaño físico (`Size`) | Requerido (`Required`) | Atributos y Comportamiento |
| :--- | :---: | :--- | :---: | :---: | :--- |
| **`Id`** | `4` | `dbLong` | 4 bytes | `False` | Autonumérico (`Counter`). Clave primaria física de la tabla. |
| **`Residencia`** | `10` | `dbText` | 20 caracteres | `True` | Texto corto. Valores almacenados: `GIJON`, `SOTO`, `OVIEDO`. |
| **`Ejercicio`** | **`4`** | **`dbLong`** | **4 bytes** | `True` | **Entero Largo (Long Integer)**. Año de la serie (ej. 2026). |
| **`UltimoNumero`** | `4` | `dbLong` | 4 bytes | `True` | Entero Largo. Cota máxima asignada en el ejercicio. |

#### Índices de la tabla `ContadorFacturas`
1. **`idx_CF_ResEjer`**:
   - Campos: `(Residencia, Ejercicio)`
   - `Primary`: `False`
   - `Unique`: `True`
   - Función: Garantiza que no exista más de un contador por residencia y ejercicio.
2. **`Index_28B8F692_E7C9_49A9`**:
   - Campo: `Id`
   - `Primary`: `True`
   - `Unique`: `True`

#### Resolución definitiva de la discrepancia documental (2 bytes vs 4 bytes)
- **Hecho probado:** En el motor físico de Access (DAO), el campo `Ejercicio` es de tipo **`dbLong` (Type = 4, 4 bytes)**.
- **Origen de la confusión:** Al consultar el esquema a través del proveedor OLEDB de ADODB (`adSchemaColumns`), el tipo devuelto es `DATA_TYPE = 3`, que en la enumeración `DataTypeEnum` de ADO corresponde a `adInteger` (entero de 4 bytes con signo). En la enumeración de DAO, el código `3` corresponde a `dbInteger` (entero de 2 bytes). Un informe preliminar interpretó el `3` de OLEDB como el `dbInteger` de DAO, deduciendo erróneamente que ocupaba 2 bytes. **Queda certificado que el campo físico es y ha sido siempre de 4 bytes (`dbLong` / `adInteger`).**

#### Registros actuales en `ContadorFacturas`
- `Id = 1` | `Residencia = GIJON` | `Ejercicio = 2026` | `UltimoNumero = 192`
- `Id = 2` | `Residencia = SOTO` | `Ejercicio = 2026` | `UltimoNumero = 20`
- `Id = 3` | `Residencia = OVIEDO` | `Ejercicio = 2026` | `UltimoNumero = 1032`

### 5.2. Procedimientos de asignación y rectificación de facturas
El flujo de numeración de facturas está implementado en `AsignarNumFactura.bas` y `modDatabase.bas`:

1. **Disparo automático (`AsignarFacturaPorColumna`):**
   - Disparado al marcar "PAGADO" en la columna de pago (`AA` en GIJÓN/SOTO, `AB` en OVIEDO).
   - Comprueba que la resolución sea válida (`SI`, `CONCEDIDA`, `REEVALUADA`).
   - Propone el siguiente número consultando:
     ```sql
     SELECT UltimoNumero FROM ContadorFacturas WHERE Residencia = '...' AND Ejercicio = Year(Now)
     ```
   - **Constatación esencial:** El código utiliza **`Year(Now)`** (el año de la fecha actual del sistema), y **no utiliza `Year(FechaSalida)`**.

2. **Incremento atómico en Access (`modDatabase.ObtenerSiguienteNumFactura`, líneas 2678-2775):**
   - Inicia transacción ADODB (`cn.BeginTrans`).
   - Ejecuta:
     ```sql
     SELECT MAX(NumFactura) FROM Ordenes WHERE Residencia = '...'
     ```
     *(Nótese que esta consulta busca el máximo global en toda la tabla `Ordenes`, sin filtrar por `Ejercicio` ni año de salida).*
   - Si `MAX(NumFactura) > UltimoNumero`, alinea `UltimoNumero = MAX(NumFactura)`.
   - Incrementa atómicamente:
     ```sql
     UPDATE ContadorFacturas SET UltimoNumero = UltimoNumero + 1 WHERE Residencia = '...' AND Ejercicio = ejercicio
     ```
   - Lee el nuevo valor y ejecuta `cn.CommitTrans`.

3. **Modificación manual de factura (`ProcesarCambioManualFactura`, líneas 557-641):**
   - Si el usuario escribe manualmente un número en la Columna C:
     - Valida duplicados locales en la hoja.
     - Ejecuta `ExisteFacturaEnBD` para comprobar que ningún otro registro en `Ordenes` tenga ese número:
       ```sql
       SELECT COUNT(*) FROM Ordenes WHERE Residencia = '...' AND NumFactura = ... AND NumOrden <> ...
       ```
     - Si no colisiona, persiste en Access (`modDatabase.ActualizarOrden numOrden, "NumFactura", nuevoValor`).
     - Invoca `SincronizarContadorFacturaManual`: si el número introducido supera el `UltimoNumero` del contador, lo adelanta; si es menor (rellenando un hueco), no modifica el contador.

4. **Borrado manual del número de factura:**
   - Si el usuario borra el contenido de la celda de la Columna C:
     - `ProcesarCambioManualFactura` ejecuta `modDatabase.ActualizarOrden CLng(numOrden), "NumFactura", "", resCode`.
     - **No decrementa ni altera `ContadorFacturas.UltimoNumero`.**
     - Ese número de factura queda como un **hueco irreversible** en el contador de la base de datos.
     - **No existe ningún registro de facturas anuladas o liberadas.**

5. **Modificación posterior de la `FechaSalida` (Columna M / Columna 13):**
   - En `Hoja2.cls`, `Hoja20.cls` y `Hoja24.cls`, la edición de la celda `FechaSalida`:
     - Invoca `VerificarDisponibilidad...`.
     - Invoca `SincronizarCambioTargetHaciaAccess`, actualizando el campo `FechaSalida` en Access.
     - **No realiza ninguna comprobación sobre el campo `NumFactura`**. Si el cambio de fecha traslada el expediente a otro año natural, la factura permanece asociada con el número y serie originales, sin emitir aviso ni forzar rectificación.

6. **Facturación diferida offline (SICO-N2):**
   - En contingencia offline, asigna la etiqueta `"PROFORMA"` en la celda y encola la operación `ASIGNAR_FACTURA_DIFERIDA`.
   - Al restaurar la conexión, `ConsolidarFacturaDiferidaFIFO` (`ModuloSeguridadLocal.bas`, líneas 1122-1188):
     - Determina el año como `anioEjercicio = Year(CDate(fechaHoraLocal))` o `Year(Now)`.
     - Invoca `modDatabase.ObtenerSiguienteNumFactura(resCode, anioEjercicio)`.
     - Sobrescribe la celda de Excel y actualiza el campo `NumFactura` en Access.

### 5.3. Inexistencia de registros históricos o tablas auxiliares de facturación
Se ha examinado la lista completa de las 11 tablas de usuario presentes en `Residencia_BE.accdb`:
1. `BDAS_Historico_Snapshot`
2. `BloqueosHabitaciones`
3. `CodigosPostales`
4. `ConfigTarifasCeldas`
5. `ContadorFacturas`
6. `EstadoHabitacionesCalendario`
7. `ListaNegra`
8. `LogActividad`
9. `Ordenes`
10. `Ordenes_Historico`
11. `Usuarios`

- **Resultado definitivo:** **No existe ninguna tabla de registro histórico de facturas** (tipo `RegistroNumerosFactura`, `FacturasEmitidas` o `HistorialFacturas`).
- El número de factura asignado reside única y exclusivamente como un valor escalar en el campo `NumFactura` de la tabla `Ordenes` (o de `Ordenes_Historico` si ha sido archivado). Si un expediente es rectificado o eliminado, no queda rastro en la base de datos de qué factura le había sido adjudicada previamente.

---

## 6. CUADRO RESUMEN DE HECHOS ACREDITADOS, AUSENCIAS Y CONCLUSIONES NO DEMOSTRADAS

| Aspecto auditado | Estado en el sistema actual | Evidencia técnica acreditada | Impacto funcional / Riesgo detectado |
| :--- | :--- | :--- | :--- |
| **Procedimiento `CrearOrden`** | **Inexistente** | 0 coincidencias en la totalidad del proyecto VBA. | No existe una puerta única de creación; la inserción está fragmentada en 5 procedimientos. |
| **Identidad permanente (`IdInstanciaOrden`)** | **Inexistente** | El campo `Id` de Access es autonumérico interno y jamás se traslada a Excel ni a la cola offline. | El sistema carece de identidad de instancia. Si `NumOrden` se reutiliza, se pierde la distinción entre expedientes. |
| **Token `colUUID` en `SYS_COLA_SYNC`** | **Identificador de transacción** | Generado por `GenerarUUID` en `ModuloSeguridadLocal.bas`. | Es efímero y exclusivo de la cola FIFO. No sirve como identificador permanente del expediente. |
| **Columnas de reparto de coste y marcas** | **No persistidas en BD** | Columna 30 (`AD`) en GIJÓN/SOTO, y Columnas 29-31 (`AC:AE`) en OVIEDO no se leen ni se guardan en Access. | Datos volátiles que se destruyen al archivar, borrar o refrescar la hoja desde Access. |
| **Esquema físico de `ContadorFacturas.Ejercicio`** | **`dbLong` (4 bytes)** | DAO Type = 4 (`dbLong`), Size = 4 bytes. | Queda resuelta la contradicción documental: la tabla física soporta enteros largos de 4 bytes. |
| **Cálculo del ejercicio de facturación** | **Basado en `Year(Now)`** | `AsignarFacturaPorColumna` pasa `Year(Now)` a `ObtenerSiguienteNumFactura`. | No respeta la regla de negocio arquitectónica de calcular el ejercicio mediante `Year(FechaSalida)`. |
| **Efecto de alterar `FechaSalida`** | **Ninguno sobre factura** | `Hoja2.cls`, `Hoja20.cls`, `Hoja24.cls` no tienen código que evalúe facturas al cambiar fecha salida. | Puede haber facturas adjudicadas en un ejercicio cuyo servicio finalice en un ejercicio diferente. |
| **Registro permanente de facturas** | **Inexistente** | Solo existe `ContadorFacturas` (que guarda el valor máximo) y el campo `NumFactura` en `Ordenes`. | La anulación o rectificación deja huecos permanentes sin trazabilidad de facturas emitidas. |

---

## 7. DECLARACIÓN FORMAL DE NO ALTERACIÓN Y RESPETO DE RESTRICCIONES

El inspector técnico certifica solemnemente que:
1. Todas las actuaciones se han limitado estrictamente a la lectura estática de archivos de código y consultas de inspección de solo lectura sobre el libro Excel y la base de datos Access.
2. No se ha ejecutado ninguna sentencia DDL ni DML de modificación sobre `Residencia_BE.accdb`.
3. No se ha modificado ninguna línea de código del libro `BDAS_Multiusuario.xlsm` ni de sus módulos asociados.
4. No se han propuesto diseños de arquitectura, esquemas alternativos ni fragmentos de código ejecutable, limitando el informe a los hechos empíricamente demostrados.
5. Se ha respetado de manera escrupulosa la regla estricta de nombres con el literal `GIJÓN`.
6. En cumplimiento de las directrices RGPD, no se ha extraído ni expuesto ningún dato de carácter personal de las personas administradas.
