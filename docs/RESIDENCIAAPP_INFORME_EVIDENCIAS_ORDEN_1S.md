# RESIDENCIAAPP — INFORME DE EVIDENCIAS Y DEPENDENCIAS DE IDENTIDAD Y SNAPSHOT (ORDEN 1S)

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Diseño arquitectónico de `IdInstanciaOrden`, numeración, facturación y archivado  
**Naturaleza:** Extracción de evidencias y consultas de solo lectura  
**Destinatario:** Antigravity (inspector técnico)  
**Dirección arquitectónica:** ChatGPT  
**Fecha de emisión:** 19 de noviembre de 2026  
**Documento entregable:** `RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md`

---

## 1. CONTROL DE INTEGRIDAD Y TRAZABILIDAD DE FUENTES AUTORIZADAS

En estricto cumplimiento del protocolo establecido en la Sección 2 de la Orden 1S, se ha efectuado el cálculo y verificación de firmas criptográficas SHA-256 de las fuentes autorizadas antes de iniciar cualquier comprobación.

### 1.1. Tabla de Hashes y Verificación Criptográfica

| Fuente | Ruta Absoluta | Tamaño (bytes) | SHA-256 Esperado | SHA-256 Actual | Estado |
| :--- | :--- | :---: | :---: | :---: | :---: |
| **Libro Activo** | [`BDAS_Multiusuario.xlsm`](file:///H:/ResidenciaApp/bdas-multiusuario/BDAS_Multiusuario.xlsm) | 6.621.963 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` | **COINCIDE (100%)** |
| **Módulo Base de Datos** | [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) | 298.794 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` | **COINCIDE (100%)** |
| **Backend de Pruebas** | [`Residencia_BE.accdb`](file:///H:/ResidenciaBD/Residencia_BE.accdb) | 7.639.040 | `853eb4e736c02934668f0cc10d5a4d393f3cd66ceb40bf887b8993115f0ed430` | `9617f99795b8266c2abb91b479bf64b9ece449ad15885ce40edbe62639dbbee9` | **DISCREPANCIA (Auditada)** |

### 1.2. Auditoría Forense del Backend de Pruebas

Se realizó una inspección diferencial completa de `Residencia_BE.accdb` para acreditar el estado de los datos:
1. **Recuentos en tablas de negocio:**
   - `Ordenes`: exactamente **2.127** registros (100% idéntico a Órdenes 1Q, 1Q-A y 1R).
   - `Ordenes_Historico`: exactamente **17** registros (100% idéntico).
   - `ContadorFacturas`: exactamente **3** registros (100% idéntico).
   - `Usuarios`: exactamente **3** registros (100% idéntico).
   - `ConfigTarifasCeldas`: exactamente **291** registros (100% idéntico).
   - `CodigosPostales`: exactamente **14.602** registros (100% idéntico).
   - `BDAS_Historico_Snapshot`: exactamente **2.118** registros (100% idéntico).
2. **Causa raíz física:** La tabla de log de auditoría `LogActividad` contiene 8.930 registros (dos registros adicionales respecto a la Orden 1R: `Id 8929` correspondiente a LOGIN del usuario `admin` a las 23:22:41 del 18/11/2026, e `Id 8930` correspondiente a LOGOUT normal a las 00:28:25 del 19/11/2026).
3. **Conclusión de integridad:** La totalidad de las tablas de datos de expedientes, tarifas, usuarios, contadores y esquemas estructurales permanecen **rigurosa y estrictamente intactas**.

---

## 2. COMPROBACIÓN A — DEPENDENCIAS REALES DEL CAMPO `Id` EN `Ordenes` Y `Ordenes_Historico`

### 2.1. Propiedades Físicas del Campo `Id` en el Esquema de Access (DAO)

| Propiedad | Tabla `Ordenes` | Tabla `Ordenes_Historico` |
| :--- | :--- | :--- |
| **Nombre del Campo** | `Id` | `Id` |
| **Posición Ordinal** | Columna 1 | Columna 1 |
| **Tipo DAO** | `dbLong` (Código 4) | `dbLong` (Código 4) |
| **Tamaño Físico** | 4 bytes (Entero largo con signo) | 4 bytes (Entero largo con signo) |
| **Autonumérico (`AutoNumber`)** | **True** (Atributo `dbAutoIncrField`, +16) | **True** (Atributo `dbAutoIncrField`, +16) |
| **Propiedad Required** | False (gestionado por el generador secuencial) | False |
| **Permite Longitud Cero** | False | False |
| **Clave Primaria (Primary Key)** | **SÍ**: Índice `Index_50734CA2_0946_4D32` (Unique=True, Primary=True) | **NO**: Ninguna clave primaria definida |
| **Índices Asociados** | Clave Primaria única sobre `Id` | **Cero índices sobre `Id`** (único índice es `idx_Hist_Res_NumOrden`) |

---

### 2.2. Inventario Exhaustivo de Sentencias SQL que Utilizan el Campo `Id`

Se ha realizado un análisis léxico y sintáctico sobre la totalidad de los 173 módulos VBA del proyecto. De las 56 sentencias SQL que interactúan con `Ordenes` y `Ordenes_Historico`, el resultado respecto al campo `Id` es el siguiente:

#### 1. Sentencias `SELECT`
- **Referencias explícitas a `Id`:** Se localiza **exactamente 1 sentencia SELECT** en todo el código base que menciona el campo `Id`:
  - **Módulo:** [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3279)
  - **Procedimiento:** `Public Sub SincronizarHojaDesdeAccess(ByVal residencia As String, ByVal hoja As Worksheet, Optional ByVal filaInicio As Long = 2)` ([líneas 3235-3410](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3235-L3410)).
  - **Línea 3279:**
    ```vba
    sql = "SELECT * FROM Ordenes WHERE Residencia = '" & EscaparSQL(residencia) & "' ORDER BY Id ASC"
    ```
  - **Uso fáctico del campo:** El campo `Id` se utiliza **exclusivamente como criterio de ordenación ascendente (`ORDER BY Id ASC`)** para volcar las filas activas en el libro Excel en el orden secuencial en que fueron dadas de alta en Access.
  - **Comportamiento en Excel:** El valor de `rs("Id")` **NUNCA se escribe en ninguna celda de Excel ni se expone al usuario**. La columna A de Excel recibe `rs("NumOrden")` ([línea 3351](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3351)). En Excel no existe ninguna columna para `Id`.
- **Resto de sentencias `SELECT`:** Ninguna otra consulta SELECT del sistema incluye el campo `Id` (seleccionan explícitamente `NumOrden, FechaPeticion, DNI...` o utilizan `SELECT TOP 1 *` para validación de estructura sin leer `Id`).

#### 2. Sentencias `INSERT`
- **Altas e inserciones ordinarias:** En `modDatabase.bas` ([L904](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L904), [L1359](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1359), [L1442](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1442), [L1725](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1725)) y en `modMigracion.bas` ([L148](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modMigracion.bas#L148)), todas las sentencias `INSERT INTO Ordenes (...)` **omiten intencionadamente el campo `Id`**, delegando su generación automática en el motor autonumérico de Access.
- **Inserciones de archivado (`INSERT INTO ... SELECT *`):** Se localizan en tres procedimientos de `modDatabase.bas`:
  1. `EliminarOrdenBD` ([línea 4381](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4381)):
     ```vba
     cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & _
                " AND Residencia = '" & EscaparSQL(claveRes) & "' AND [UsuarioModificacion] = '" & EscaparSQL(tokenActual) & "'"
     ```
  2. `EliminarOrdenesLoteBD` ([línea 4517](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4517)):
     ```vba
     cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes " & _
                "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND NumOrden IN (" & inClause & ") AND ((DNI IS NOT NULL...) OR (Nombre IS NOT NULL...))"
     ```
  3. `EliminarTodasOrdenesSedeBD` ([línea 4631](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4631)):
     ```vba
     cn.Execute "INSERT INTO Ordenes_Historico SELECT * FROM Ordenes " & _
                "WHERE Residencia = '" & EscaparSQL(claveRes) & "' AND ((DNI IS NOT NULL...) OR (Nombre IS NOT NULL...))"
     ```

#### 3. Sentencias `UPDATE`
- **Resultado:** **0 sentencias UPDATE utilizan el campo `Id`**.
- La totalidad de actualizaciones del sistema ([`ActualizarOrden`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L2487), Lease Locks en [L1791](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1791), [L4298](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4298)) utilizan invariablemente la cláusula `WHERE NumOrden = ... AND Residencia = '...'`.

#### 4. Sentencias `DELETE`
- **Resultado:** **0 sentencias DELETE utilizan el campo `Id`**.
- La totalidad de eliminaciones físicas ([L4378](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4378), [L4389](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4389), [L4513](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4513), [L4523](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4523), [L4626](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4626), [L4637](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4637), [L4986](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4986)) filtran por `(Residencia, NumOrden)` o por `FechaSalida <= ...`.

---

### 2.3. Procedimientos que Reciben o Devuelven Valores `Id`

- **Resultado Acreditado:** **CERO procedimientos**.
- No existe ninguna función, subrutina ni propiedad en todo el proyecto VBA que acepte un parámetro denominado `Id` (o referido al identificador de la orden) ni que retorne un valor `Id`.
- *(Nota técnica: La única rutina del proyecto con un parámetro `id` es `IdentificadorEnLista(ByVal id As String, ByVal arr As Variant)` en `ModuloLavanderia_Gijon.bas`, correspondiente a cadenas de texto de códigos de lavandería de prendas).*

---

### 2.4. Operaciones de Archivado y Preservación de `Id`

Al ejecutarse `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes WHERE ...`:
1. **Comportamiento físico del motor Access (Jet/ACE):** Puesto que `SELECT *` proyecta las 43 columnas en su orden ordinal físico, el valor del campo `Ordenes.Id` (columna 1) es transferido explícitamente a la columna `Ordenes_Historico.Id` (columna 1). En Jet/ACE, un campo autonumérico permite la inserción explícita de valores numéricos procedentes de otro origen autonumérico.
2. **Consecuencia inmediata:** El registro copiado a `Ordenes_Historico` **conserva el mismo valor de `Id`** que poseía en `Ordenes`.
3. **Evidencia empírica en producción:**
   - En la auditoría criptográfica, el recuento de solapamientos entre `Ordenes.Id` y `Ordenes_Historico.Id` es **exactamente CERO** (0 IDs compartidos).
   - Los IDs de `Ordenes_Historico` oscilan entre `4308` y `4342`.
   - Cuando una orden histórica fue archivada, su `Id` técnico pasó a `Ordenes_Historico`. Al borrarse la fila activa mediante `DELETE FROM Ordenes WHERE NumOrden = ...`, ese `Id` desapareció de `Ordenes`.
   - Si con posterioridad se reutilizó el mismo `NumOrden` (algoritmo ORD-10), la nueva orden en `Ordenes` recibió un **nuevo autonumérico `Id` completamente diferente** (por ejemplo, en GIJÓN 641, el registro histórico posee `Id = 4314` y el registro activo posterior posee `Id = 4319`).
   - Sin embargo, al volver a archivarse la nueva versión activa, el código ejecuta un `DELETE FROM Ordenes_Historico WHERE NumOrden = ...` previo, lo que **destruye físicamente la fila histórica previa y su `Id` original**.

---

### 2.5. Dependencias en Importación, Sincronización, Borrado y Consultas Relacionales

- **Importación (`modMigracion.bas`):** Completamente desacoplada de `Id`. Se generan nuevos números autonuméricos en Access.
- **Sincronización (`SincronizarFilaHaciaAccess` / `SincronizarHojaDesdeAccess`):** Desacoplada de `Id`. `SincronizarHojaDesdeAccess` sólo usa `Id` para ordenar el cursor (`ORDER BY Id ASC`).
- **Borrado (`EliminarOrdenBD` y variantes):** Desacoplado de `Id`. Opera sobre `(Residencia, NumOrden)`.
- **Consultas Relacionales (JOINs / Claves Foráneas):** **CERO dependencias**. No existe ningún `JOIN`, subconsulta relacional ni restricción de integridad referencial en toda la base de datos que utilice `Id` para vincular tablas. La vinculación entre tablas en el sistema se realiza exclusivamente mediante `(Residencia, NumOrden)`.

---

### 2.6. Cuadro Comparativo: Identificador Técnico (`Id`) vs. Identificador de Negocio (`NumOrden`)

| Criterio | Identificador Técnico (`Id`) | Identificador de Negocio (`NumOrden`) |
| :--- | :--- | :--- |
| **Definición** | Entero largo autonumérico interno de Access (`AutoNumber`, 4 bytes). | Número entero de expediente en la residencia (`Long`, 4 bytes). |
| **Presencia en Excel** | **Inexistente** (0 columnas, 0 celdas, 0 referencias en hojas). | **Columna A** de las hojas de residencia (GIJÓN, SOTO, OVIEDO). |
| **Clave en Access** | Clave Primaria en `Ordenes` (sin PK en `Ordenes_Historico`). | Parte de la clave única compuesta `(Residencia, NumOrden)`. |
| **Uso en Lógica VBA** | Únicamente en `ORDER BY Id ASC` en `modDatabase.bas:3279`. | En todas las operaciones de CRUD, bloqueos, Lease Lock y sincronización. |
| **Estabilidad** | Autoincremental en inserción; mutable si se reinserta. | Asignado al huésped; reutilizado históricamente bajo ORD-10. |

---

## 3. COMPROBACIÓN B — CICLO DE VIDA COMPLETO DE `BDAS_Historico_Snapshot`

### 3.1. Esquema Físico Completo en Access (DAO)

- **Nombre de la Tabla:** `BDAS_Historico_Snapshot`
- **Recuento de Registros en Producción:** **2.118 registros**.
- **Número Total de Campos:** **41 campos**.
- **Clave Primaria:** **NINGUNA** (0 claves primarias definidas).
- **Índices Físicos:** **NINGUNO** (0 índices definidos en la colección `Indexes`).

#### Catálogo Completo de Campos y Propiedades Físicas:

| Ordinal | Nombre del Campo | Tipo DAO | Tamaño (bytes) | Required | AutoNumber |
| :---: | :--- | :--- | :---: | :---: | :---: |
| 1 | `Id` | `dbLong` | 4 | False | **True** |
| 2 | `NumOrden` | `dbLong` | 4 | False | False |
| 3 | `FechaPeticion` | `dbDate` | 8 | False | False |
| 4 | `NumFactura` | `dbLong` | 4 | False | False |
| 5 | `Finalidad` | `dbText` | 50 | False | False |
| 6 | `Empleo` | `dbText` | 50 | False | False |
| 7 | `Situacion` | `dbText` | 50 | False | False |
| 8 | `Evaluacion` | `dbText` | 50 | False | False |
| 9 | `Turno` | `dbText` | 50 | False | False |
| 10 | `Comision` | `dbText` | 50 | False | False |
| 11 | `DNI` | `dbText` | 20 | False | False |
| 12 | `Rango` | `dbText` | 50 | False | False |
| 13 | `Nombre` | `dbText` | 100 | False | False |
| 14 | `Apellidos` | `dbText` | 100 | False | False |
| 15 | `FechaEntrada` | `dbDate` | 8 | False | False |
| 16 | `FechaSalida` | `dbDate` | 8 | False | False |
| 17 | `DiasUso` | `dbInteger` | 2 | False | False |
| 18 | `PAX` | `dbInteger` | 2 | False | False |
| 19 | `Resolucion` | `dbText` | 50 | False | False |
| 20 | `NumHabIndividuales` | `dbInteger` | 2 | False | False |
| 21 | `NumHabDobles` | `dbInteger` | 2 | False | False |
| 22 | `CamaSuple` | `dbInteger` | 2 | False | False |
| 23 | `HabitacionesAsignadas` | `dbText` | 100 | False | False |
| 24 | `DtoFamNum` | `dbText` | 50 | False | False |
| 25 | `Importe` | `dbCurrency` | 8 | False | False |
| 26 | `EstadoPago` | `dbText` | 50 | False | False |
| 27 | `PagadoResidencia` | `dbCurrency` | 8 | False | False |
| 28 | `Observaciones` | `dbMemo` | 0 | False | False |
| 29 | `Telefono` | `dbText` | 30 | False | False |
| 30 | `Direccion` | `dbText` | 150 | False | False |
| 31 | `CodigoPostal` | `dbText` | 10 | False | False |
| 32 | `Poblacion` | `dbText` | 100 | False | False |
| 33 | `Provincia` | `dbText` | 50 | False | False |
| 34 | `Email` | `dbText` | 100 | False | False |
| 35 | `ConsentimientoRGPD` | `dbBoolean` | 1 | False | False |
| 36 | `FechaGrabacion` | `dbDate` | 8 | False | False |
| 37 | `Residencia` | `dbText` | 20 | False | False |
| 38 | `Solapamiento` | `dbBoolean` | 1 | False | False |
| 39 | `TipoHuesped` | `dbText` | 50 | False | False |
| 40 | `FechaCreacion` | `dbDate` | 8 | False | False |
| 41 | `UsuarioCreacion` | `dbText` | 50 | False | False |

---

### 3.2. Procedimientos que Interactúan con `BDAS_Historico_Snapshot`

Se ha analizado todo el código fuente del proyecto para identificar las operaciones que acceden a esta tabla:

| Tipo Operación | Sentencia SQL | Módulo | Procedimiento | Líneas |
| :--- | :--- | :--- | :--- | :---: |
| **INSERT** | *Inexistente en código VBA* | — | — | — |
| **UPDATE** | *Inexistente en código VBA* | — | — | — |
| **SELECT** | `SELECT NumOrden, FechaPeticion, NumFactura... FROM BDAS_Historico_Snapshot WHERE ...` | [`ModuloArchivadoDatosBDAS.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloArchivadoDatosBDAS.bas#L380) | `buscarEnBDAS` | [L376-381](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloArchivadoDatosBDAS.bas#L376-L381) |
| **DELETE** | `DELETE FROM BDAS_Historico_Snapshot WHERE FechaSalida <= DateAdd('yyyy', -anosRetencion, Date())` | [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4991) | `PurgarHistoricoRGPD` | [L4973-5003](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4973-L5003) |

---

### 3.3. Llamadores Directos e Indirectos

1. **`buscarEnBDAS`:**
   - **Llamador directo:** [`BusquedaOrdenNombreFactura.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BusquedaOrdenNombreFactura.bas#L152) en la subrutina `BuscarValorEnColumna` ([línea 152](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BusquedaOrdenNombreFactura.bas#L152)).
   - **Función en el sistema:** Proporciona los resultados de búsqueda cuando el operador ejecuta la búsqueda global por Orden, Factura, DNI o Nombre desde el formulario o atajo del Ribbon, combinando mediante `UNION ALL` los registros de `Ordenes_Historico` y `BDAS_Historico_Snapshot`.
2. **`PurgarHistoricoRGPD`:**
   - **Llamador directo:** [`BorradoFilasHojasBDAS.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoFilasHojasBDAS.bas#L39) en la subrutina `BorrarFilasBDAS` ([línea 39](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoFilasHojasBDAS.bas#L39)) y el formulario de gestión RGPD.
   - **Función en el sistema:** Purga física y definitiva de registros personales con antigüedad superior a 3 años según el RD 933/2021.

---

### 3.4. Momento del Ciclo de Vida en el que se Genera y se Utiliza

1. **Origen y Naturaleza:** La tabla `BDAS_Historico_Snapshot` no se genera ni se alimenta durante la operativa ordinaria de la aplicación. Corresponde a una **instantánea estática de respaldo y seguridad** generada durante la importación inicial del histórico preexistente de la residencia (contiene 2.118 registros correspondientes a peticiones registradas entre el 02/12/2025 y el 21/09/2026, con salidas hasta el 07/10/2026).
2. **Desconexión del Ciclo Ordinario de CRUD:**
   - La creación de reservas (`modDatabase.CrearOrden`) no escribe en `BDAS_Historico_Snapshot`.
   - La edición de celdas y sincronización (`SincronizarFilaHaciaAccess`) no escribe en `BDAS_Historico_Snapshot`.
   - El archivado ordinario (`EliminarOrdenBD`, `EliminarOrdenesLoteBD`, `ArchivarOrdenesAntiguasBD`) escribe **únicamente en `Ordenes_Historico`**, nunca en `BDAS_Historico_Snapshot`.
3. **Uso en Producción:** Se utiliza exclusivamente como **repositorio pasivo de consulta histórica** en búsquedas globales (`buscarEnBDAS`) y como tabla sujeta a purga de retención legal en `PurgarHistoricoRGPD`.

---

### 3.5. Correspondencia de Campos con `Ordenes` y `Ordenes_Historico`

- **Campos en `Ordenes`:** 43 campos.
- **Campos en `Ordenes_Historico`:** 43 campos.
- **Campos en `BDAS_Historico_Snapshot`:** 41 campos.
- **Diferencia Exacta:** `BDAS_Historico_Snapshot` **carece exactamente de 2 campos**:
  - `UsuarioModificacion` (dbText, 50 caracteres).
  - `UltimaModificacion` (dbDate, 8 bytes).
- **Razón técnica acreditada:** Los campos `UsuarioModificacion` y `UltimaModificacion` fueron incorporados en la arquitectura V5.2 para el soporte del protocolo Lease Lock en las tablas activas e históricas operativas. La instantánea snapshot, al ser estática y previa a dicha fase, conserva la estructura original de 41 columnas.
- **Propiedad Required:** En `BDAS_Historico_Snapshot`, los campos `NumOrden` y `Residencia` tienen `Required = False` (al igual que en `Ordenes_Historico`), a diferencia de `Ordenes` donde son obligatorios (`Required = True`).

---

### 3.6. Dependencias Funcionales ante la Incorporación de `IdInstanciaOrden`, `AnioOrden` y `AnioFactura`

1. **Sentencia UNION en `buscarEnBDAS`:** En [`ModuloArchivadoDatosBDAS.bas:370-381`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloArchivadoDatosBDAS.bas#L370-L381), la consulta SQL nombra explícitamente las columnas seleccionadas (`NumOrden, FechaPeticion, NumFactura, DNI, NomComp, Telefono, Direccion, FechaGrabacion, Residencia, FechaSalida, HabitacionesAsignadas, EstadoPago`). Por tanto:
   - La simple adición de nuevas columnas (`IdInstanciaOrden`, `AnioOrden`, `AnioFactura`) a `Ordenes` u `Ordenes_Historico` **no romperá sintácticamente la consulta de búsqueda actual**, ya que esta no utiliza `SELECT *`.
   - Sin embargo, si en el futuro se desea proyectar o filtrar en búsquedas históricas por `IdInstanciaOrden` o por ejercicio mediante `UNION ALL`, `BDAS_Historico_Snapshot` carecerá de dichos campos salvo que sea adaptada en el esquema.
2. **Purga RGPD:** `PurgarHistoricoRGPD` ejecuta un `DELETE FROM BDAS_Historico_Snapshot WHERE FechaSalida <= ...`, por lo que es inmune a la adición de nuevas columnas en la tabla.

---

## 4. COMPROBACIÓN C — GUID COMPLETO Y FUNCIONAMIENTO OFFLINE

### 4.1. Declaraciones Actuales de `CoCreateGuid` en el Proyecto

La API nativa de Windows `CoCreateGuid` (`ole32.dll`) se encuentra declarada con soporte dual de 32/64 bits en [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4-L17):

```vba
Private Type GUID_TYPE
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID_TYPE) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function CoCreateGuid Lib "ole32.dll" (ByRef pguid As GUID_TYPE) As Long
#End If
```

---

### 4.2. Procedimiento Actual `ObtenerTokenSesionInstancia` y Truncamiento a 48 Bits

En [`modDatabase.bas:1029-1059`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1029-L1059):

```vba
If CoCreateGuid(uGuid) = 0 Then
    hexGuid = Right$("00000000" & Hex$(uGuid.Data1), 8) & _
              Right$("0000" & Hex$(uGuid.Data2 And &HFFFF&), 4)
Else
    ...
End If
m_SessionTokenCache = pc & "#" & UCase$(hexGuid)
```

- **Estructura Factual de la Reducción:**
  - `uGuid.Data1` (Long = 32 bits): genera 8 dígitos hexadecimales.
  - `uGuid.Data2` (Integer = 16 bits): genera 4 dígitos hexadecimales.
  - **Bits conservados:** $32 + 16 = 48$ bits (12 caracteres hexadecimales).
  - **Bits descartados:** `Data3` (16 bits) y `Data4(0 To 7)` (64 bits) son completamente ignorados, descartándose 80 de los 128 bits del GUID estándar RFC 4122.
- **Motivo Arquitectónico del Truncamiento:** El campo `UsuarioModificacion` en Access tiene un límite físico estricto de `dbText(50)`. Para dar cabida al formato `Usuario@NombrePC#HexGUID` sin desbordar los 50 caracteres, se redujo el GUID a 12 caracteres hexadecimales.

---

### 4.3. Otras Funciones de Generación de Identificadores en el Proyecto

En [`ModuloSeguridadLocal.bas:1625-1628`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1625-L1628) existe la función:

```vba
Private Function GenerarUUID(ByVal res As String, ByVal seq As Long) As String
    Randomize
    GenerarUUID = Left$(res, 3) & "-" & Format$(Now, "yyyymmddhhnnss") & "-" & _
                  Format$(seq, "0000") & "-" & Right$("0000" & Hex$(Int(Rnd() * 65535)), 4)
End Function
```

- **Naturaleza:** Es un identificador sintético basado en concatenación textual (`RES-AAAAMMDDHHMMSS-SECUENCIAL-HEXALEATORIO`).
- **Uso exclusivo:** Se utiliza únicamente en la columna L (`colUUID`) de la hoja de cola `SYS_COLA_SYNC` como token de idempotencia interna de la transacción de sincronización. No constituye un GUID binario de 128 bits ni se envía a Access.

---

### 4.4. Tipos y Formatos de Almacenamiento Disponibles en Access (DAO)

Para almacenar un identificador GUID completo de 128 bits, el motor de base de datos Access (ACE/Jet) ofrece técnicamente las siguientes opciones en su esquema físico:

1. **`dbGUID` (Código DAO 15):**
   - Denominado en Access *"Id. de réplica"* (Replication ID / GUID nativo).
   - Tamaño: exactamente **16 bytes** binarios.
   - Representación textual nativa en consultas y formularios: `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` (38 caracteres con llaves y guiones).
   - Ventaja física: almacenamiento óptimo de 16 bytes e indexación nativa de alta velocidad.
2. **`dbText` (Código DAO 10):**
   - Cadena de texto de tamaño fijo o variable.
   - Tamaño requerido:
     - 36 caracteres: `XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX` (formato estándar RFC 4122 canónico sin llaves).
     - 38 caracteres: `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}` (formato canónico de Windows con llaves).
     - 32 caracteres: `XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX` (representación hexadecimal pura sin guiones).
3. **`dbBinary` (Código DAO 9):**
   - Campo binario puro de 16 bytes.

*(Nota: En cumplimiento de la Restricción Absoluta de la Sección 6, Antigravity no selecciona ni propone el tipo de campo; expone exclusivamente las capacidades nativas verificadas en el motor).*

---

### 4.5. Serialización de Datos en `SYS_COLA_SYNC` y Archivos WAL

El subsistema SICO-N2 implementa un búfer dual no volátil para garantizar la continuidad operativa offline:

#### 1. Hoja Oculta `SYS_COLA_SYNC`
Estructura de 13 columnas definida en el enum `eColCola` ([`ModuloSeguridadLocal.bas:38-52`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L38-L52)):
- Columna 1 (`colIDTransaccion`): Secuencial numérico Long (1, 2, 3...).
- Columna 2 (`colFechaHoraLocal`): Cadena `yyyy-mm-dd hh:nn:ss`.
- Columna 3 (`colUsuario`): Nombre de usuario operador.
- Columna 4 (`colResidencia`): Clave de sede (`GIJON`, `SOTO`, `OVIEDO`).
- Columna 5 (`colNumOrden`): Nº ORDEN (provisional 900xxx o definitivo).
- Columna 6 (`colTipoOperacion`): `UPDATE_CAMPO`, `ASIGNAR_FACTURA_DIFERIDA`, `PRERESERVA`, `ELIMINAR_ORDEN`, `LOG_OFFLINE`.
- Columna 7 (`colCampoModificado`): Nombre del campo (`EstadoPago`, `HabitacionesAsignadas`, etc.).
- Columna 8 (`colValorAnterior`): Pre-image auditada.
- Columna 9 (`colValorNuevo`): Post-image.
- Columna 10 (`colEstadoSync`): `PENDIENTE`, `SINCRONIZANDO`, `CONCILIADO`, `CONFLICTO_HAB`, `ERROR`.
- Columna 11 (`colReintentos`): Contador de reintentos.
- Columna 12 (`colUUID`): Token de idempotencia de la transacción local.
- Columna 13 (`colCRC`): Suma de verificación CRC16 de integridad de la fila.

#### 2. Diario de Disco WAL (`.wal`)
- **Ubicación:** `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\SYS_COLA_SYNC_<COMPUTERNAME>.wal` ([`L1583`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1583)).
- **Formato:** Archivo plano de texto con codificación estándar, delimitado por barras verticales (`|`), donde los caracteres reservados son escapados mediante `SanitizarPipe` ([`L1485`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1485), [`L1617`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1617)).
- **Comportamiento ante fallo o cierre abrupto:** `RehidratarColaDesdeDiscoWALSiNecesario` ([`L1499-1569`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1499-L1569)) lee el archivo WAL al abrir el libro y rehidrata en un solo bloque matricial la hoja `SYS_COLA_SYNC`.

---

### 4.6. Ciclo de Conciliación y Procedimiento `ResolverNumOrdenProvisionalEnBD`

Cuando una reserva se crea offline:
1. **Asignación provisional:** Se le asigna un número provisional en el rango `900001..999999` ([`L598-615`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L598-L615)).
2. **Encolado de operaciones:** La operación de alta (`PRERESERVA`) y las ediciones subsiguientes sobre esa reserva se encolan asociadas a ese número provisional.
3. **Reconciliación al volver la red (`SincronizarColaOfflineHaciaAccess`):**
   - Al procesar la fila de la cola ([`L1000-1003`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1000-L1003)), se invoca `ResolverNumOrdenProvisionalEnBD`:
     ```vba
     sqlMax = "SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '" & modDatabase.EscaparSQL(resCode) & "' AND NumOrden < 900001"
     vMax = modDatabase.ExecuteScalar(sqlMax)
     nuevoNum = CLng(vMax) + 1
     ```
   - Actualiza la celda en la hoja Excel con el `nuevoNum` definitivo ([`L1387`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1387)).
   - **Reescritura de la cola pendiente:** Itera sobre el resto de filas de `SYS_COLA_SYNC` reemplazando el número provisional por el definitivo ([`L1391-1396`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1391-L1396)).
   - Invoca `modDatabase.SincronizarFilaHaciaAccess` para volcar los datos a la tabla `Ordenes` de Access.

---

### 4.7. Viabilidad Fáctica de Preservación de un GUID Completo

A partir de la inspección técnica de los componentes reales, se acredita que:
1. **Generación:** Win32 `CoCreateGuid` ya está integrado y operativo en el binario. Formatear los 128 bits completos en una cadena canónica RFC 4122 (ej. `36` caracteres) requiere únicamente leer `Data1`, `Data2`, `Data3` y `Data4(0..7)` mediante las funciones nativas de VBA.
2. **Persistencia en Excel:** Una cadena de 36 caracteres puede almacenarse de forma nativa en cualquier columna de la hoja Excel o en una estructura en memoria.
3. **Persistencia en `SYS_COLA_SYNC` y WAL:** Las rutinas de encolado (`EncolarOperacionOffline`), volcado a disco (`VolcarColaHaciaDiscoWAL`) y rehidratación (`RehidratarColaDesdeDiscoWALSiNecesario`) manejan cadenas delimitadas por `|` sin restricciones de longitud. Un GUID no contiene caracteres de control ni barras verticales, por lo que su serialización y deserialización es 100% transparente.
4. **Independencia del Número de Orden:** Un identificador GUID generado en el momento exacto de crear la reserva permanece **completamente inmutable**, independientemente de si el `NumOrden` asignado inicialmente es provisional (`900001`) y se renombra posteriormente al consolidarse con Access (`645`).

---

## 5. CUADRO RESUMEN DE HECHOS ACREDITADOS, AUSENCIAS Y CUESTIONES NO VERIFICABLES

| Aspecto | Hecho Acreditado | Tipo | Evidencia en Código / Esquema |
| :--- | :--- | :---: | :--- |
| **Dependencia de `Ordenes.Id`** | Sólo se referencia en un `ORDER BY Id ASC` en la carga a Excel. | Positivo | [`modDatabase.bas#L3279`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L3279). |
| **Paso de parámetros `Id`** | Cero procedimientos reciben o retornan el campo `Id` de las órdenes. | **Negativo (Ausencia)** | Escaneo integral de firmas procedimentales. |
| **Relaciones / Claves Foráneas con `Id`** | Cero relaciones, JOINs o claves foráneas físicas utilizan `Id`. | **Negativo (Ausencia)** | Catálogo ADOX y consultas SQL. |
| **Copia de `Id` al Histórico** | `INSERT INTO ... SELECT *` copia el valor de `Id`, pero se pierde ante borrados previos por `NumOrden`. | Positivo | [`modDatabase.bas#L4381`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4381), [`L4517`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4517), [`L4631`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4631). |
| **Esquema de `BDAS_Historico_Snapshot`** | Posee 41 campos, carece de `UsuarioModificacion` y `UltimaModificacion`, y tiene 0 índices y 0 PKs. | Positivo | Definición física DAO en `TableDefs`. |
| **Escritura en `BDAS_Historico_Snapshot`** | Cero procedimientos en el código activo ejecutan INSERT o UPDATE sobre esta tabla. | **Negativo (Ausencia)** | Escaneo integral de sentencias SQL. |
| **Lectura en `BDAS_Historico_Snapshot`** | Sólo la consulta `UNION ALL` de búsqueda textual `buscarEnBDAS` la lee. | Positivo | [`ModuloArchivadoDatosBDAS.bas#L376-381`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloArchivadoDatosBDAS.bas#L376-L381). |
| **Purga en `BDAS_Historico_Snapshot`** | Sólo `PurgarHistoricoRGPD` ejecuta DELETE sobre ella por fecha de salida. | Positivo | [`modDatabase.bas#L4991`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4991). |
| **Generador GUID Win32** | `CoCreateGuid` está declarado y operativo en `ole32.dll`. | Positivo | [`modDatabase.bas#L11-17`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L11-L17). |
| **Truncamiento Actual de GUID** | Actualmente se descartan 80 bits del GUID para acotar el token a 12 hex en `UsuarioModificacion`. | Positivo | [`modDatabase.bas#L1042-1051`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1042-L1051). |
| **UUID en Cola Offline** | Es un string sintético (`RES-Fecha-Seq-Rnd`) usado solo como idempotencia de la transacción local. | Positivo | [`ModuloSeguridadLocal.bas#L1625-1628`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1625-L1628). |
| **Viabilidad de GUID Completo Offline** | La arquitectura de `SYS_COLA_SYNC` y WAL soporta cadenas arbitrarias y mantiene identidad inmutable. | Positivo | Análisis de `ModuloSeguridadLocal.bas`. |

---

## 6. DECLARACIÓN FORMAL DE NO ALTERACIÓN Y RESPETO DE RESTRICCIONES

1. **Inspección de Solo Lectura:** La totalidad de las comprobaciones se ha ejecutado mediante inspección estática de código fuente y consultas de solo lectura en DAO/ADODB.
2. **Inalterabilidad de Archivos y Registros:** No se ha modificado ningún módulo VBA, celda de Excel, tabla ni registro de la base de datos de producción o pruebas.
3. **Protección de Datos Personales (RGPD):** No se han publicado ni extraído datos identificativos personales de huéspedes.
4. **Regla Canónica de Nombres:** Se ha utilizado estrictamente el texto literal `GIJÓN` conforme a las reglas permanentes del proyecto.
5. **Abstención de Diseño Arquitectónico:** Antigravity actúa única y exclusivamente como inspector y extractor técnico de evidencias; no formula propuestas de implementación, selección de tipos de campo, pseudocódigo ni decisiones de arquitectura, las cuales quedan reservadas a la Dirección Arquitectónica ejercida por ChatGPT.

---
*Informe técnico cerrado y verificado por Antigravity a 19 de noviembre de 2026.*
