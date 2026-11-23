# ResidenciaApp — Especificación arquitectónica v1.0

**Estado:** definitiva como diseño arquitectónico; pendiente de tres verificaciones de entorno antes de autorizar la implementación.  
**Ámbito:** Excel 2016 (VBA) + Access (ACE/DAO 32 bits) + funcionamiento multiusuario y offline.  
**Fecha de cierre de diseño:** 19 de noviembre de 2026.  
**Base documental:** diseño v0.2, órdenes de auditoría 1R y 1S–1X y decisiones funcionales expresas del usuario.  
**Restricción esencial:** este documento no autoriza modificar producción, el esquema Access ni el código VBA. La implementación comienza únicamente después de superar las puertas de salida del apartado 14.

---

## 0. Cómo interpretar esta especificación

Se utilizan cuatro estados:

- **Regla funcional confirmada:** indicada expresamente por el usuario o ya acordada en el proyecto.
- **Decisión arquitectónica adoptada:** solución elegida por la Arquitecta por coherencia e integridad del sistema. No se debe sustituir por otra sin evidencia técnica o una decisión funcional posterior.
- **Pendiente de comprobación en laboratorio/entorno:** el diseño está definido, pero una propiedad concreta debe probarse con Excel 2016 y ACE/DAO de 32 bits o extraerse del entorno operativo.
- **Bloqueo de implantación:** no se autoriza modificar código o esquema hasta resolverlo.

Las auditorías aportadas se consideran evidencia para contrastar, no órdenes de implementación. Las soluciones de cada informe se han evaluado y solo se incorporan cuando se consideran adecuadas.

## 1. Resolución ejecutiva

La arquitectura objetivo separa tres conceptos que hoy están mezclados:

1. **Identidad del expediente:** `IdInstanciaOrden`, estable durante toda la vida del expediente.
2. **Clave de negocio de la orden:** `(Residencia, AnioOrden, NumOrden)`, única entre órdenes activas.
3. **Consumo de factura:** `(Residencia, AnioFactura, NumFactura)`, registrado de forma permanente y no reutilizable.

El alta conserva el flujo real de los operadores: primero introducen el Nº ORDEN en A y después la Fecha de Petición en B. A inicia un borrador con identidad permanente; B aporta el ejercicio y permite solicitar la reserva definitiva a Access.

La factura tiene una serie independiente de la orden. `AnioFactura` procede de `FechaSalida`, no de `FechaPeticion` ni del reloj del equipo. El negocio ha confirmado que **no existen facturas agrupadas: una factura física pertenece a un único Nº ORDEN**.

El histórico se identifica por `IdInstanciaOrden`, no por el número reutilizable. Los cinco campos funcionales que antes solo existían localmente en Excel pasarán a persistirse en Access y en el histórico. El usuario ha confirmado que sus valores desaparecen al cerrar el libro y nunca se han conservado de forma permanente: por tanto, **no se realizará backfill de esos cinco campos; los nuevos campos comenzarán en NULL**.

La arquitectura queda cerrada como diseño. No está autorizado aún el cambio físico porque faltan dos comprobaciones del entorno real: universo de facturas positivas en `BDAS_Historico_Snapshot` y estado efectivo de la cola/WAL en los equipos operadores. Los dominios funcionales y tipos preferentes de los cinco campos quedan definidos por las aclaraciones del usuario; PREFLIGHT solo debe comprobar que las validaciones existentes del libro no los contradicen.

## 2. Identidad permanente y ubicación Excel

### 2.1 `IdInstanciaOrden`

**Decisión arquitectónica adoptada:** cada expediente recibe un GUID completo de 128 bits, almacenado en formato de texto canónico de 36 caracteres:

`XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX`

- Mayúsculas y sin llaves.
- Tipo Access preferente: `dbText(36)`.
- Una instancia nueva recibe un GUID nuevo, aunque reutilice residencia, ejercicio y número de un expediente archivado.
- El mismo GUID se conserva en Excel, Access, histórico y cola offline.
- `Ordenes.Id` permanece como autonumérico técnico local de Access; no es una identidad compartida entre tablas.
- El identificador de sesión/lease y el identificador de operación no se confunden con el GUID del expediente.

`dbText(36)` sustituye a `dbGUID` como opción principal. Su compatibilidad concreta con DAO, SQL, índices, VBA y WAL se verificará en laboratorio antes de crear el campo físico.

### 2.2 Ubicación en Excel: columna AJ

**Ubicación preferente adoptada:** `AJ` (columna 36), oculta, en GIJÓN, SOTO y OVIEDO.

AJ se acepta como ubicación de diseño, condicionada a estas garantías:

- Toda ordenación permitida debe mover el registro completo, al menos `A:AJ`; ninguna ordenación parcial puede dejar inmóviles `AD`, `AC:AE` o `AJ`.
- Los vaciados completos deben limpiar AJ expresamente.
- La sincronización Excel↔Access y la cola offline deben transportar AJ y los cinco campos funcionales.
- El saneamiento de columnas excedentes debe continuar comenzando en `AK` (37), sin invadir AJ.
- AJ debe seguir oculta y no alterar la presentación ordinaria.

La ampliación efectiva del AutoFilter hasta AJ es la primera alternativa que se probará, no una implementación ya aprobada. En GIJÓN y SOTO el rango persistido actual termina en AB; en OVIEDO el filtro se aplica por código y no se encontró un `_FilterDatabase` persistido equivalente. Si ampliar A:AJ crea flechas no deseadas o no garantiza la ordenación completa, se deberá usar ordenación controlada de registro completo y bloquear las vías de ordenación parcial.

### 2.3 Token de sesión/lease

**Decisión:** mantener el token actual de 48 bits derivado de `CoCreateGuid` para identificar la instancia propietaria de una sesión/lease.

- No se amplía `UsuarioModificacion` (actualmente `dbText(50)`) a 80 caracteres.
- No se sustituye el token por `NombrePC#GUID36`.
- Los 48 bits identifican una sesión de trabajo en la red local; no son una credencial criptográfica ni sustituyen autenticación.
- `IdInstanciaOrden` y `IdOperacion` utilizan GUID completos independientes.

Solo se reabrirá esta decisión ante evidencia reproducible de colisión o un requisito adicional concreto.

## 3. Alta y ciclo de vida de órdenes: flujo A → B

### 3.1 Al introducir A (Nº ORDEN)

1. Comprobar que se trata de una fila nueva y que no tiene GUID.
2. Generar una sola vez `IdInstanciaOrden` y guardarlo en AJ.
3. Mantener el número introducido como candidato pendiente, no como reserva anual confirmada.
4. Crear/actualizar un borrador local durable asociado al GUID.
5. No calcular `AnioOrden` con `Year(Now)`, no asumir el ejercicio actual y no validar definitivamente la unicidad anual mientras B esté vacía.

La edición posterior de cualquier dato de la fila —DNI, nombre, fechas, resolución, campos funcionales, etc.— sigue perteneciendo al borrador hasta que Access confirme la reserva.

### 3.2 Al introducir B (Fecha de Petición)

1. Validar que B contiene una fecha válida.
2. Calcular `AnioOrden = Year(FechaPeticion)`.
3. Crear un `IdOperacion` propio para el intento de reserva, distinto de `IdInstanciaOrden`.
4. Enviar a Access residencia, ejercicio, número candidato, GUID y el conjunto coherente de datos introducidos.
5. Access valida y reserva `(Residencia, AnioOrden, NumOrden)` en una única operación transaccional.
6. Solo después de una confirmación inequívoca de Access la orden pasa a estado confirmado y las ediciones siguientes pueden seguir el flujo ordinario de actualización.

Si la clave ya está ocupada, Access devuelve conflicto. La aplicación puede ofrecer el siguiente número libre de la misma residencia y ejercicio. Si se acepta otro número, se genera un nuevo `IdOperacion` de intento y se conserva el mismo `IdInstanciaOrden`.

### 3.3 Borrador, cambios y refresco

- Mientras no exista reserva confirmada, ningún cambio en `C:AE` ni en otras columnas de datos lanza `UPDATE` contra Access.
- Una fecha escrita en B no basta para habilitar actualizaciones: se espera confirmación de la reserva.
- Tampoco se permite emitir, asignar o consumir una factura ni generar un documento fiscal definitivo mientras el expediente siga en borrador o la reserva no haya sido confirmada. Las marcas de pago y demás cambios permanecen locales hasta entonces.
- El borrador debe sobrevivir a `SincronizarHojaDesdeAccess`, a un refresco completo y al cierre/reapertura del libro.
- No se confiará en que el libro XLSM se guarde para persistir el borrador. La arquitectura actual marca `ThisWorkbook.Saved = True` al cerrar, y los cinco campos locales se pierden al cerrar.
- La persistencia recuperable del borrador debe residir en el WAL local o en un mecanismo equivalente de escritura durable, con CRC, versión de formato y confirmación de escritura antes de informar al operador de que el dato está protegido.
- El libro reconstruido combinará las órdenes confirmadas de Access con los borradores y operaciones pendientes locales, asociándolos por GUID.

Si el libro se cierra después de escribir A pero antes de B, el borrador se recuperará del registro local durable. Si ese registro no pudo escribirse, la aplicación debe informar de que el borrador no está protegido; nunca prometer recuperación no garantizada.

### 3.4 Alta offline y números provisionales

Hay que distinguir dos casos:

- **Número manual introducido por el operador:** es un candidato y no se cambia silenciosamente al recuperar conexión. Si colisiona para el ejercicio de B, se informa y se solicita aceptación de otro número.
- **Número provisional asignado por el sistema en contingencia:** se mantiene la estrategia actual de números provisionales `900xxx` solo después de adaptarla al GUID y a `AnioOrden`. La conciliación sustituye el número provisional por el definitivo, conservando el mismo GUID.

En ambos casos, los reintentos reutilizan la identidad de la orden y el identificador de la misma operación pendiente. No se crea un segundo expediente por pérdida de respuesta después del `Commit`.

### 3.5 Cambios posteriores de A o B

Una vez confirmada la orden:

- Modificar A no genera un GUID nuevo; cambia la clave de negocio del mismo expediente y exige validar unicidad por GUID y transacción.
- Modificar B dentro del mismo ejercicio conserva `AnioOrden` y actualiza `FechaPeticion`.
- Modificar B a otro ejercicio recalcula `AnioOrden` y valida la nueva clave. Si existe conflicto, se rechaza sin dejar cambios parciales.
- Ninguna actualización parcial puede alterar A o B sin ejecutar estas reglas.

## 4. Modelo de datos de órdenes

### 4.1 `Ordenes` activa

Modelo lógico y físico preferente:

| Campo | Tipo lógico/físico preferente | Regla |
|---|---|---|
| `Id` | `dbLong` AutoNumber actual | Identificador técnico Access, no identidad del expediente |
| `IdInstanciaOrden` | `dbText(36)` | GUID canónico; único y obligatorio en toda orden nueva confirmada |
| `AnioOrden` | `dbLong` | `Year(FechaPeticion)`; obligatorio en toda orden nueva confirmada |
| `AnioFactura` | `dbLong`, nullable | `Year(FechaSalida)` si existe factura vigente; NULL si no hay factura vigente |
| `NumOrden` | `dbLong` actual | Número visible; candidato hasta confirmar alta |
| `NumFactura` | `dbLong` actual, nullable | Factura vigente; NULL cuando no existe asociación vigente |
| Cinco campos funcionales | `REPARTIR COSTE`: `dbText(2)` con dominio `NO`/`SI`; `TAG` y `R`: `dbInteger` (2 bytes), enteros de hasta 4 dígitos | Datos persistentes, inicialmente NULL en registros existentes |

Nombres físicos ASCII obligatorios: `AnioOrden` y `AnioFactura`; no se utilizarán `AñoOrden` ni `AñoFactura` como nombres SQL/DAO. En interfaz se pueden mostrar textos con ñ.

Índices objetivo:

- Único sobre `IdInstanciaOrden`.
- Único sobre `(Residencia, AnioOrden, NumOrden)` en órdenes activas.
- Índices adicionales solo si los patrones reales de consulta lo justifican.

El índice único actual `(Residencia, NumOrden)` es incompatible con numeración anual reutilizable y deberá sustituirse de forma controlada, después de tratar filas antiguas sin `FechaPeticion` y colisiones potenciales.

### 4.2 `Ordenes_Historico`

- Mantener el autonumérico `Id` como identificador técnico independiente.
- Incorporar `IdInstanciaOrden`, `AnioOrden`, `AnioFactura` y los cinco campos funcionales.
- Índice único por `IdInstanciaOrden`.
- La clave de negocio `(Residencia, AnioOrden, NumOrden)` no será única en histórico, porque dos expedientes distintos pueden haber reutilizado el mismo número y ejercicio.
- Utilizar listas explícitas de campos en INSERT/UPDATE. Quedan descartadas las inserciones históricas basadas en `INSERT ... SELECT *`.

### 4.3 Política histórica por GUID

**Decisión arquitectónica adoptada:** la unidad histórica es `IdInstanciaOrden`, no `(Residencia, AnioOrden, NumOrden)`.

La regla anterior de una versión por clave de negocio no puede coexistir con la reutilización del número dentro del mismo ejercicio y con la obligación de conservar íntegros ambos expedientes.

Ejemplo: un expediente `G1` y otro `G2` pueden haber utilizado GIJÓN/2027/150 en momentos distintos. Ambos deben conservar su identidad e histórico. Al archivar G2 no se sobrescribe G1.

Se mantiene la regla de una única mejor versión elegible por GUID:

- Si no hay histórico válido previo, se inserta el nuevo.
- Si el histórico previo es inválido y el nuevo es válido, se sustituye.
- Si ambos son válidos, la última versión elegible sustituye a la anterior, aunque tenga menos columnas cumplimentadas que la anterior, siempre que ella misma cumpla el umbral y todos los requisitos.
- Un registro nuevo no elegible no sustituye ni elimina un histórico válido.
- Un borrado definitivo que no archive no modifica el histórico de otro GUID con la misma clave visible.

Para los registros antiguos sin GUID, la migración no inferirá identidad solo por residencia y número. Las coincidencias activo/histórico se vincularán únicamente con evidencia suficiente de que son el mismo expediente; los casos dudosos se separarán y se informarán como excepciones.

## 5. Histórico, archivo y borrado

### 5.1 Recuento de columnas y umbral

Los recuentos correctos de columnas de datos del usuario son:

| Residencia | Intervalos | Total |
|---|---|---:|
| GIJÓN | `A:T` (20) + `V:AA` (6) + `AD` (1) | **27** |
| SOTO | `A:P` (16) + `S:T` (2) + `V:AA` (6) + `AD` (1) | **25** |
| OVIEDO | `A:U` (21) + `W:AE` (9) | **30** |

El umbral para archivo es **al menos 17 columnas de datos cumplimentadas**. Las fórmulas quedan excluidas. Los cinco campos funcionales cuentan como columnas de datos; AJ y demás columnas técnicas no cuentan.

### 5.2 Nueve campos obligatorios

- A: Nº ORDEN.
- B: Fecha de Petición.
- I: DNI.
- K: Nombre.
- L: Fecha de Entrada.
- M: Fecha de Salida.
- P: Resolución.
- Y: Población.
- Dirección: W en GIJÓN/SOTO; X en OVIEDO.

### 5.3 Coherencia exacta entre factura y resolución (C/P)

Solo se consideran discrepancias los dos supuestos aprobados:

- **Tipo 1:** C contiene una factura y P pertenece a `{NO, DENEGADA, DESESTIMADA, RENUNCIA}`.
- **Tipo 2:** C está vacía y P pertenece a `{SI, CONCEDIDA, REEVALUADA}`.

Los demás estados no constituyen discrepancia por esta regla. No se endurece la condición exigiendo que toda factura tenga siempre una resolución positiva ni que toda resolución positiva tenga factura si el caso no pertenece a los conjuntos anteriores.

### 5.4 Matriz de decisión de borrado/archivo

Se evalúa en este orden:

1. **Menos de 17 columnas de datos:** borrado definitivo sin preguntar, independientemente de campos obligatorios o discrepancia C/P.
2. **Al menos 17, pero solo 0–6 de los nueve campos obligatorios presentes:** borrado definitivo sin preguntar.
3. **Al menos 17 y 7 u 8 campos obligatorios presentes:**
   - Sin discrepancia C/P: preguntar entre borrado definitivo y modificar/cancelar conservando la fila activa.
   - Con discrepancia C/P: borrado definitivo.
4. **Al menos 17 y los 9 campos obligatorios presentes:**
   - Con discrepancia C/P: advertir y abortar toda la operación; conservar la fila activa.
   - Sin discrepancia: elegible para archivar.

### 5.5 Atomicidad del archivo

Para una fila elegible:

1. Localizar y validar el expediente por `IdInstanciaOrden`.
2. Evaluar la matriz sobre los datos coherentes del expediente.
3. Insertar/actualizar la versión histórica de ese GUID.
4. Eliminar la fila activa por GUID.
5. Confirmar ambas modificaciones dentro de una sola transacción Access.

Los procesos individual, lote y archivo completo deben aplicar la misma matriz y el mismo contrato. No debe haber reglas divergentes entre los distintos caminos.

El registro permanente de consumo de facturas no se elimina por archivar, borrar o purgar la fila del expediente. La política RGPD debe permitir conservar la clave mínima necesaria para impedir la reutilización; los datos personales y el vínculo identificativo se limitarán al mínimo legalmente permitido.

## 6. Modelo de facturación

### 6.1 Reglas de negocio cerradas

- Serie independiente por residencia y ejercicio.
- `AnioFactura = Year(FechaSalida)` (columna M).
- Nunca usar `Year(Now)`, `FechaPeticion` ni la hora de encolado como fuente del ejercicio.
- Un expediente tiene como máximo una factura vigente asociada.
- Una factura física corresponde a un único expediente: **no existen facturas agrupadas**.
- Un expediente puede tener varios consumos históricos a lo largo de su vida por corrección, rectificación o reemisión.
- Un número consumido nunca se libera aunque se borre su asociación vigente o se archive/elimine el expediente.
- Valor ausente de factura: `NumFactura = NULL`; los nuevos registros no utilizarán cero como sustituto de ausencia.

### 6.2 Tabla única `RegistroNumerosFactura`

**Decisión arquitectónica adoptada:** una sola tabla de registro permanente; no se crea una estructura N:M ni una tabla intermedia de factura compartida.

Cada fila representa un número físico consumido. Campos lógicos requeridos:

| Campo lógico | Tipo preferente | Función |
|---|---|---|
| `Id` | `dbLong` AutoNumber | Clave técnica de la fila |
| `Residencia` | Texto ASCII corto compatible con el catálogo existente | Sede emisora |
| `AnioFactura` | `dbLong` | Ejercicio fiscal de la factura |
| `NumFactura` | `dbLong` | Número físico consumido |
| `IdInstanciaOrden` | `dbText(36)` | GUID del único expediente asociado; no único en la tabla |
| `IdOperacion` | `dbText(36)` | GUID de la operación que consumió el número; único |
| `FechaConsumo` | `dbDate` | Marca temporal de la operación, para auditoría |
| `Origen/EstadoConciliacion` | Texto/estado corto | Distinguir altas nuevas de importación histórica pendiente de vincular |

Índices:

- **Único:** `(Residencia, AnioFactura, NumFactura)`.
- **Único:** `IdOperacion`.
- **No único:** `IdInstanciaOrden`.

El campo `IdInstanciaOrden` podrá quedar NULL únicamente en filas de consumo heredado todavía no vinculadas de forma inequívoca; esas filas deberán marcarse como pendientes de conciliación y bloquearán la reutilización de su tupla. En nuevas emisiones será obligatorio y deberá corresponder a un único expediente. Al terminar la migración no se admitirán consumos nuevos sin GUID.

No se creará un índice único por `IdInstanciaOrden`: la tabla debe conservar varias facturas históricas consumidas por el mismo expediente a lo largo del tiempo.

### 6.3 Uso de `ContadorFacturas`

`ContadorFacturas` sigue siendo un generador de secuencia, no el registro de consumos.

- `ContadorFacturas.Ejercicio` está verificado por DAO en la Orden 1T como `dbLong`, Type 4, 4 bytes. La discrepancia `dbLong`/`dbInteger` queda cerrada para el backend inspeccionado; no debe reaparecer como decisión pendiente.
- El índice único `(Residencia, Ejercicio)` existente se mantiene como base de unicidad del contador.
- Debe crearse de forma segura la fila de un nuevo ejercicio si no existe.
- La inicialización/reconciliación del contador tomará en cuenta los consumos existentes del mismo ejercicio, sin asumir que `UltimoNumero` por sí solo representa todos los consumos.
- Los huecos en una serie no hacen reutilizable un número que esté en `RegistroNumerosFactura`.
- La asignación manual debe pasar por la misma tabla y unicidad. Si se usa un número manual libre por debajo del máximo, se registra su consumo; si supera el máximo, el contador debe actualizarse para no proponer luego un número inferior ya adelantado.

### 6.4 Transacción de emisión

Automática, manual y FIFO deben llamar a un único servicio transaccional de Access que:

1. Verifica el expediente por GUID y su estado.
2. Valida `FechaSalida` y calcula `AnioFactura`.
3. Obtiene/bloquea o crea el contador de `(Residencia, AnioFactura)`.
4. Selecciona/reserva el número según la modalidad.
5. Inserta el consumo en `RegistroNumerosFactura` con `IdOperacion` único.
6. Actualiza `NumFactura` y `AnioFactura` vigentes en `Ordenes`.
7. Confirma contador, ledger y asociación vigente con un único `Commit`.

Si `IdOperacion` ya existe, Access devuelve el resultado original, sin consumir otro número. Si el mismo `IdOperacion` llega con un payload incompatible, se rechaza como reutilización indebida de clave de idempotencia.

La hoja Excel, los resúmenes y el documento de factura solo se actualizan como confirmados después del `Commit` de Access. Si Access no actualiza exactamente un expediente, el llamador no debe continuar ni escribir una confirmación local.

### 6.5 Corrección de `FechaSalida` con factura emitida

**Decisión arquitectónica adoptada:** el cambio ordinario de `FechaSalida` a otro ejercicio queda bloqueado cuando existe una factura vigente. Se requiere una operación explícita de corrección/reasignación.

- Si el cambio mantiene el mismo ejercicio, la asociación puede mantenerse si las demás reglas lo permiten.
- Si cambia el ejercicio, no se reescribe el año de la factura consumida.
- La factura antigua permanece registrada como consumida.
- La operación explícita cierra la asociación vigente anterior y, cuando corresponda, emite otra factura del nuevo ejercicio con un nuevo `IdOperacion`.
- Si se borra `FechaSalida` y existe factura vigente, debe exigirse una regularización; no se permite una factura vigente sin ejercicio determinable.

Esta regla protege la coherencia del historial fiscal y evita que una edición ordinaria de fecha cambie de forma silenciosa una asignación ya emitida.

### 6.6 Datos heredados de facturación

La Orden 1R documenta en el backend de pruebas:

- `Ordenes`: 2.127 filas; 653 con `NumFactura > 0`; 4 con `NumFactura = 0`.
- `Ordenes_Historico`: 17 filas; una con `NumFactura > 0`.
- Siete números repetidos en OVIEDO/ejercicio de salida 2026 (`10, 15, 16, 180, 182, 183, 1031`) en 15 expedientes distintos.
- Los cuatro ceros son excepciones, no facturas válidas.

La confirmación del usuario de que no existen facturas agrupadas implica que esos siete grupos no representan una regla de negocio. Son excepciones heredadas que deben resolverse antes de crear el índice único. Solo hay colisión si se repite la tupla completa `(Residencia, AnioFactura, NumFactura)`; el mismo número puede aparecer en otra sede o ejercicio.

No se sanearán ni renumerarán esos registros automáticamente. Los ceros se reportarán como excepción y no se importarán como consumos positivos. Los consumos positivos de activo, histórico y snapshot se conciliarán por tupla completa y evidencia de expediente.

## 7. Los cinco campos funcionales

| Residencia | Columna | Campo |
|---|---:|---|
| GIJÓN | AD | REPARTIR COSTE (`dbText(2)`, `NO`/`SI`) |
| SOTO | AD | REPARTIR COSTE (`dbText(2)`, `NO`/`SI`) |
| OVIEDO | AC | REPARTIR COSTE (`dbText(2)`, `NO`/`SI`) |
| OVIEDO | AD | TAG (`dbInteger`, 2 bytes) |
| OVIEDO | AE | R (`dbInteger`, 2 bytes) |

**Decisión funcional confirmada:** los cinco campos deben persistir en Access, histórico, refrescos de Excel y operaciones offline; cuentan para el umbral de 17 columnas.

El usuario ha confirmado que, al cerrar el libro, esos valores desaparecen y nunca se han conservado de forma permanente.

Consecuencias:

- No se desarrolla backfill de los cinco campos.
- Los nuevos campos Access se inicializan en NULL para los registros activos e históricos existentes.
- No se intenta reconstruir valores desde libros antiguos ni seleccionar una copia de Excel como fuente autoritativa.
- Desde la implantación, cada modificación queda persistida por GUID.
- Si durante el corte existe un libro abierto con valores transitorios en memoria, no se consideran una fuente histórica respaldada. No se importarán automáticamente.
- `REPARTIR COSTE` admite exclusivamente `NO` o `SI`: tipo físico preferente `dbText(2)`, nullable. La aplicación y Access deben validar el dominio, normalizando espacios y mayúsculas antes de persistir.
- `TAG` y `R` contienen números enteros de un máximo de cuatro dígitos: tipo físico preferente `dbInteger` (DAO Type 3, 2 bytes), nullable. No usar texto ni memo, no aceptar decimales ni valores que superen cuatro dígitos. Un valor vacío se almacena como `NULL`; un cero introducido explícitamente no se debe convertir en `NULL`.
- `dbInteger` admite sobradamente cuatro cifras. La validación de negocio debe imponer el límite funcional de cuatro dígitos, independiente del rango técnico mayor del tipo Access. No se ha indicado que los ceros iniciales tengan significado; la decisión actual trata `TAG` y `R` como números, no como códigos de texto.
- La inspección PREFLIGHT comprobará que las validaciones y formatos de las celdas son compatibles con estos dominios. Si las celdas del XLSM guardado están vacías, ello no será un bloqueo ni justificará backfill.

La evidencia actual confirma que `R` en OVIEDO interviene en una validación asociada a PAX. Las columnas `AC` en GIJÓN/SOTO activan formularios de reservas. No se debe inferir de esos usos el dominio físico de los valores almacenados.

## 8. Sincronización, cola offline y concurrencia

### 8.1 Todas las operaciones técnicas se dirigen por GUID

Los UPDATE, DELETE, archivo, bloqueo, lectura de instancia y sincronización deben localizar el expediente por `IdInstanciaOrden`.

`(Residencia, AnioOrden, NumOrden)` se utiliza para validar la unicidad de negocio, no como identidad permanente para modificar una instancia. La auditoría 1W constató que el `WHERE` de `EjecutarUpdateCondicionalOrden` usa actualmente `(Residencia, NumOrden)`, aunque controla `RecordsAffected` y `ActualizarOrden` propaga el Boolean. Los llamadores de facturación que ignoran ese Boolean deben corregirse.

### 8.2 Contrato Excel → Access

- Enviar `IdInstanciaOrden` y campos explícitamente mapeados.
- Calcular y validar `AnioOrden` desde B.
- Incluir los cinco campos funcionales según residencia.
- Diferenciar borrador, reserva, modificación, factura, archivo y borrado.
- No enviar UPDATE mientras el borrador no tenga reserva Access confirmada.
- No considerar éxito ni modificar vistas/resúmenes como confirmados antes de la respuesta positiva de Access.

### 8.3 Contrato Access → Excel

- Usar listas explícitas de campos, no `SELECT *`.
- Reconstruir filas por GUID, no por posición física o número visible.
- Descargar AJ y los cinco campos funcionales.
- Preservar las columnas de fórmulas específicas de cada residencia.
- Fusionar órdenes confirmadas con borradores y operaciones locales pendientes, sin asociar datos a otro expediente.
- Limpiar los valores obsoletos de las filas que ya no existen, sin dejar valores huérfanos en AD, AC:AE o AJ.

### 8.4 Cola `SYS_COLA_SYNC` y WAL

La estructura actual documentada tiene 13 columnas (`colIDTransaccion` a `colCRC`), estados como `PENDIENTE`, `SINCRONIZANDO`, `CONCILIADO`, `CONFLICTO_HAB` y `ERROR`, un token sintético de operación en la columna L y CRC16. El WAL local utiliza líneas de texto delimitadas por `|`, con saneamiento de delimitadores y recuperación al abrir el libro.

La cola nueva debe separar:

- `IdInstanciaOrden`: identidad del expediente.
- `IdOperacion`: identidad de una petición concreta; formato GUID canónico.
- Versión del formato de cola.
- Tipo y estado de operación.
- Residencia, `AnioOrden`/`AnioFactura` cuando proceda, número y payload.
- Los cinco campos cuando formen parte de la operación.
- Revisión/orden local que permita recuperar el último borrador coherente.

Un expediente puede generar muchas operaciones; nunca se reutiliza el GUID del expediente como IdOperacion.

**Persistencia del borrador:** se escribirá un snapshot/revisión durable del borrador en el WAL, asociado al GUID. La cola no puede depender únicamente de que la hoja oculta se guarde con el libro.

**Transición desde el formato actual:** no se borran ni reinterpretan filas pendientes por defecto. Antes del cambio se debe:

1. Verificar si existen operaciones pendientes en la hoja oculta y en los WAL locales de los equipos.
2. Si es posible, vaciar/conciliar la cola antigua antes del despliegue, usando la versión antigua del programa.
3. Si quedan operaciones, migrar solo aquellas que se puedan relacionar inequívocamente con un expediente. El token UUID sintético antiguo se conserva como referencia de origen y se vincula a un nuevo IdOperacion una única vez.
4. Poner en cuarentena las operaciones ambiguas; no inventar GUID ni resolver por número sin año si hay más de una coincidencia.
5. Mantener lector compatible durante la transición cuando se necesite procesar registros antiguos; el escritor nuevo utiliza únicamente el formato versionado.

El estado actual de pendientes por estación no ha sido observado en esta conversación; está pendiente la lectura local definida en el apartado 12.

### 8.5 Idempotencia y recuperación

- El identificador de operación se genera una sola vez al crear la operación local y se reutiliza en cada reintento.
- El consumo de factura usa `IdOperacion` único en `RegistroNumerosFactura`.
- La reserva de orden se reconoce por GUID único y por la misma petición de reserva; el reintento de un GUID ya insertado debe devolver la orden existente si su payload de alta coincide, no insertar otra.
- Las actualizaciones repetibles deben ser idempotentes o registrar el resultado de operación cuando tengan efectos laterales.
- Tras un `Commit` en Access seguido de fallo al actualizar la cola, el reintento recupera el resultado ya confirmado y no vuelve a consumir una factura ni a crear otra orden.

### 8.6 Lease y actualización condicional

- El lease identifica la instancia de orden por GUID y la sesión propietaria por el token de 48 bits.
- Un cambio se confirma solo si la instancia esperada existe y el lease es válido.
- Adquisición, expiración, renovación, validación y liberación deben compartir predicados coherentes.
- `RecordsAffected <> 1`, lease perdido o GUID no encontrado generan conflicto/error. El llamador no modifica Excel como si la operación hubiera tenido éxito.

## 9. Integridad de AJ, filtros y ordenación

### 9.1 Evidencia de la Orden 1X

- `ws.Rows(...).Delete Shift:=xlUp` opera sobre la fila completa A:XFD; desplaza juntos los campos funcionales y AJ.
- No se encontraron ordenaciones ni inserciones de filas operativas programadas en VBA.
- `_FilterDatabase` de GIJÓN y SOTO termina en AB.
- OVIEDO aplica el filtro mediante `MostrarFilasPorIntervaloFechas`; no se encontró un `_FilterDatabase` persistido equivalente.
- Los vaciados completos no limpian AJ; el saneamiento de excedentes comienza en AK.

### 9.2 Regla de ordenación

**Regla obligatoria:** ninguna ordenación permitida puede operar solo sobre A:AB. Debe ordenar el registro completo al menos hasta AJ.

Primera opción a probar: extender el rango efectivo del AutoFilter hasta AJ en las tres residencias, verificando la interfaz de Excel 2016. En GIJÓN/SOTO se modificará el rango persistido; en OVIEDO se ajustará el rango generado por código. La ocultación de flechas adicionales con `VisibleDropDown:=False` es una posibilidad de prueba, no una premisa.

También se probará Datos → Ordenar. Si el operador puede iniciar una ordenación parcial fuera del AutoFilter, deberá impedirse o sustituirse por una ordenación controlada de registro completo. No se elige solución definitiva hasta superar pruebas funcionales en Excel 2016.

### 9.3 Limpieza

Añadir limpieza explícita de AJ en:

- `BorrarDatosAlojamientoOptimizado`.
- Ruta de vaciado total de `ArchivarOrdenesAntiguasBD`.

Mantener el saneamiento de columnas excedentes desde AK. Los borrados de fila completa continúan desplazando todos los datos asociados.

## 10. Datos heredados y migración

### 10.1 Órdenes sin fecha de petición

La Orden 1R documentó una fila activa sin `FechaPeticion` (OVIEDO, Nº ORDEN 2746) y 12 filas sin `FechaPeticion` en `Ordenes_Historico`.

No se puede derivar `AnioOrden` de una fecha nula. Esas filas requieren tratamiento de migración individual. No se asignará el año actual ni un año inferido por proximidad, orden de fila o `FechaSalida`.

La migración debe:

- Generar un GUID por instancia real.
- Dejar `AnioOrden` sin valor hasta resolverlo con evidencia.
- Mantener separados los registros ambiguos.
- No crear ni activar el índice único anual hasta resolver la fila activa sin ejercicio o demostrar en laboratorio una estrategia segura para registros legacy con año NULL.

### 10.2 Asignación inicial de GUID

- Cada orden activa actual recibe una identidad GUID propia.
- Cada registro histórico antiguo recibe una identidad GUID propia salvo que exista evidencia suficiente de que es la representación archivada de la misma instancia activa.
- La mera coincidencia de residencia/número/ejercicio no basta para fusionar instancias si existe posibilidad de reutilización.
- Las 9 coincidencias activo/histórico identificadas en 1R y las divergencias de factura se tratan como casos de conciliación, no como prueba automática de identidad compartida.
- La tabla histórica puede conservar distintas instancias con la misma clave visible.

### 10.3 Backfill de consumos de factura

Fuentes que deben evaluarse:

1. `Ordenes` activas.
2. `Ordenes_Historico` operativo.
3. `BDAS_Historico_Snapshot`, si contiene facturas positivas aún no representadas en las dos fuentes anteriores.

`BDAS_Historico_Snapshot` tiene 2.118 registros y 41 campos; es una instantánea histórica estática, separada del CRUD ordinario, utilizada en búsquedas y sometida a purga RGPD. Su campo `NumFactura` existe, pero el recuento positivo y el rango de ejercicios no aparecen en los informes documentales disponibles. Se mantiene sin modificar y se incluirán sus consumos positivos en el backfill solo después de extraer y conciliar sus tuplas.

Reglas:

- Importar únicamente números positivos con `AnioFactura` derivable de `FechaSalida`, salvo tratamiento excepcional documentado.
- `NumFactura = 0` se registra como excepción, no como consumo válido.
- La misma tupla completa repetida en varias fuentes puede ser la misma instancia representada en snapshot/activo/histórico o una colisión entre expedientes distintos. Se concilia con evidencia antes de asignar GUID.
- No hay facturas agrupadas; una tupla completa no puede vincularse a dos expedientes distintos.
- Una fila histórica heredada cuya instancia no pueda vincularse todavía puede mantener temporalmente `IdInstanciaOrden = NULL` y estado `PENDIENTE_LEGACY`; la tupla queda reservada y no se puede volver a emitir. La conciliación debe resolverla antes de cerrar la migración.
- El `IdOperacion` para un consumo legacy se crea una sola vez durante el backfill y se conserva en el informe/resultado de migración para que la migración sea idempotente.

No se borra ni se altera `BDAS_Historico_Snapshot` como parte del nuevo modelo de facturación. Su función histórica y purga RGPD se mantienen. `RegistroNumerosFactura` conservará la tupla mínima necesaria para impedir reutilización; el vínculo GUID y cualquier dato personal se limitarán al mínimo legalmente permitido y se revisarán frente a la política de retención.

### 10.4 Los cinco campos: sin backfill

No existe fuente persistente de valores antiguos; el usuario confirmó que desaparecen al cerrar el libro. Se crearán las columnas nuevas con NULL para los registros existentes. No se compararán copias, no se reconstruirán valores y no se elegirá una hoja como fuente de verdad.

### 10.5 `ContadorFacturas`

`ContadorFacturas.Ejercicio` es `dbLong` (DAO Type 4, 4 bytes) en la inspección directa de la Orden 1T. La discrepancia de tipo está cerrada para el backend auditado. La revisión de la copia exacta de destino será un control de despliegue, no una tarea arquitectónica abierta.

## 11. Incompatibilidades actuales que debe resolver la implementación

| Área | Situación auditada | Contrato de destino |
|---|---|---|
| Identidad | Código actual opera principalmente con `(Residencia, NumOrden)`; no usa GUID permanente | Operaciones técnicas por `IdInstanciaOrden`; clave anual solo para unicidad de negocio |
| Índice de órdenes | Único `(Residencia, NumOrden)` sin ejercicio | Único `(Residencia, AnioOrden, NumOrden)` tras migración |
| Alta | A activa lógica de reserva antes de conocer B | A crea borrador/GUID; B determina año y activa reserva |
| Actualizaciones en borrador | `Worksheet_Change` puede intentar sincronizar antes de existir orden Access | No enviar UPDATE antes de reserva confirmada; persistir cambios en WAL |
| Histórico | Borrados por clave reutilizable y `SELECT *` | Histórico por GUID, upsert por instancia y listas explícitas |
| Facturas | Secuencia usa reloj local, MAX sin ejercicio, contador separado e idempotencia insuficiente | Ledger único, ejercicio desde FechaSalida, transacción e IdOperacion |
| Facturas repetidas | Siete números repetidos en datos de prueba OVIEDO/2026; negocio confirma no agrupadas | Resolver colisiones de tupla completa antes de índice único |
| Campos funcionales | Cinco campos no se persisten ni se sincronizan | Campos Access/History/queue, valores existentes NULL |
| Cola | 13 campos, UUID sintético de operación, no porta GUID de expediente | Formato versionado con GUID de expediente e IdOperacion separados |
| Refresco | `SELECT *`, orden por Id; los campos locales no persistidos pueden desalinearse | Selección explícita y reconstrucción por GUID, preservando borradores |
| Ordenación | Filtro hasta AB; riesgo de mover solo A:AB | Toda ordenación permitida desplaza A:AJ; solución se valida en Excel 2016 |
| Limpieza | Vaciados no limpian AJ; saneamiento desde AK | ClearContents explícito en AJ y mantener saneamiento desde AK |
| Lease | Predicados de actualización y sincronización difieren | Contratos coherentes por GUID y propietario de sesión |
| Campos físicos nuevos | El usuario confirma `REPARTIR COSTE` = `NO`/`SI` y `TAG`/`R` = enteros de hasta 4 dígitos | Tipos preferentes definidos: `dbText(2)` y `dbInteger` (2 bytes); PREFLIGHT solo verifica validaciones/compatibilidad del libro |

## 12. Estado de la inspección final y comprobaciones aún necesarias

La documentación disponible permite cerrar varios aspectos sin repetir auditorías anteriores:

| Punto solicitado | Resultado documental disponible | Estado |
|---|---|---|
| Usos y persistencia de los cinco campos | 1U confirma campos editables, no mapeados a Access y algunos usos en eventos VBA. El usuario confirma que no persisten al cerrar; además aclara dominios y tipos de TAG/R. | **Cerrado funcionalmente: `REPARTIR COSTE` `dbText(2)` NO/SI; TAG/R `dbInteger` de hasta 4 dígitos. Verificar solo las validaciones de Excel.** |
| Estructura de `BDAS_Historico_Snapshot` | 1S: 2.118 filas, 41 campos, sin PK ni índices; SELECT en búsqueda, DELETE en purga, sin INSERT/UPDATE ordinario. | **Estructura y función cerradas; facturas positivas por ejercicio pendientes de extracción** |
| Estructura de `SYS_COLA_SYNC`/WAL | 1T: 13 columnas, estados, token sintético, CRC16 y formato de líneas WAL delimitadas por `|`. | **Formato conocido; recuento real de pendientes y WAL actual pendientes de extracción por equipo** |
| `ContadorFacturas.Ejercicio` | Orden 1T: DAO `dbLong`, Type 4, 4 bytes. | **Cerrado; no repetir como contradicción pendiente** |
| Conteo columnas | Intervalos recontados: 27/25/30. | **Cerrado** |
| Facturas agrupadas | Usuario confirma que no existen. | **Cerrado: relación 1 expediente–1 factura física** |
| Valores históricos de los cinco campos | Usuario confirma que nunca han persistido. | **Cerrado: no backfill; nuevos campos NULL** |

**Límite de esta inspección:** los artefactos disponibles en esta conversación/biblioteca son informes y documentos. No están disponibles aquí el backend `.accdb`, el libro `.xlsm` operativo ni los archivos WAL por estación. Por tanto, no se ha ejecutado una lectura viva de esos ficheros y no se afirma que la cola actual esté vacía ni se inventa el recuento de facturas de Snapshot.

### 12.1 Las tres comprobaciones de entorno pendientes

Se realizará una única inspección de solo lectura, acotada a:

1. **Compatibilidad de los cinco campos:** comprobar las validaciones de datos, formatos y posibles fórmulas en las celdas correspondientes, sin inferir el tipo a partir de valores guardados. El dominio funcional ya está definido: `REPARTIR COSTE` = `NO`/`SI`; `TAG` y `R` = enteros de hasta cuatro dígitos. Si el XLSM guardado está vacío, consignarlo sin crear un bloqueo. No extraer datos de huéspedes ni valores transitorios identificables.
2. **Snapshot de facturas:** recuento `NumFactura > 0`, distribución por `Residencia` y `Year(FechaSalida)`, mínimo/máximo de ejercicios, positivos sin FechaSalida y tuplas duplicadas; solo agregados y excepciones técnicas sin PII.
3. **Cola/WAL:** filas por estado y tipo en `SYS_COLA_SYNC`; existencia, número de líneas, estados, CRC válido/inválido y operaciones pendientes en los WAL locales. Debe cubrir cada estación que tenga perfil/diario propio, no solo el PC desde el que se ejecute la inspección.

No se volverán a extraer los siete duplicados, los cuatro ceros, las fechas nulas ya documentadas, las nueve coincidencias activo/histórico ni los índices ya verificados por 1R, salvo que la inspección detecte contradicción concreta.

## 13. Contratos de las operaciones centrales

Los nombres aquí describen responsabilidades; las firmas VBA reales se definirán después de inspeccionar los módulos y dependencias vigentes.

### 13.1 Reserva de orden

Entrada: GUID, IdOperacion de reserva, residencia, número candidato, FechaPeticion, AnioOrden y snapshot de datos del borrador.

Garantías: ejercicio derivado de B; unicidad por sede/año/número; repetición del mismo GUID/petición no crea otra orden; conflicto devuelve resultado explícito; cambiar número tras conflicto conserva GUID y usa nuevo IdOperacion.

### 13.2 Actualización

Entrada: GUID, operación, campos permitidos, control de lease y versión/estado esperado.

Garantías: WHERE por GUID; cambios A/B revalidan clave anual; ningún UPDATE antes de reserva confirmada; `RecordsAffected` exacto; la hoja no presenta un éxito hasta confirmación Access.

### 13.3 Factura

Entrada: GUID, IdOperacion, modalidad automática/manual y número solicitado si es manual.

Garantías: `AnioFactura` desde FechaSalida; unicidad de tupla; alta del ledger, contador y asociación vigente en una transacción; mismo IdOperacion devuelve mismo resultado; reemisión con nueva operación conserva consumo previo.

### 13.4 Archivo/borrado

Entrada: GUID y ámbito individual/lote/total.

Garantías: misma matriz; upsert histórico por GUID; borrado activo transaccional; un borrado no elegible no toca otro GUID; consumo de factura no se elimina.

## 14. Puertas de salida y autorización de implementación

### Puerta 0 — Inspección documental y funcional

**Completada:** decisiones de numeración, identidad, no agrupación, campos sin backfill, histórico por GUID, coherencia C/P y regla de factura con cambio de año están definidas en esta v1.0.

### Puerta 1 — Inspección viva de solo lectura

**Pendiente:** completar las tres comprobaciones del apartado 12.1 en el entorno de Antigravity, con hashes de archivos y evidencia agregada sin PII.

### Puerta 2 — Laboratorio técnico

Antes de modificar código de producción:

- Probar `dbText(36)` en ACE/DAO 32 bits: INSERT, SELECT, UPDATE, comparación, índices y SQL.
- Probar `AnioOrden`/`AnioFactura` `dbLong` y comportamiento de NULL legacy.
- Verificar que las validaciones y formatos existentes de los cinco campos no contradicen sus dominios ya definidos; no volver a decidir sus tipos por inferencia.
- Probar AutoFilter/ordenación A:AJ en Excel 2016 en las tres hojas, incluyendo Datos → Ordenar.
- Probar limpieza AJ en vaciado manual y archivo total; confirmar que AK+ se mantiene.
- Probar formato WAL versionado, recuperación, CRC e idempotencia.
- Resolver en laboratorio las filas sin FechaPeticion, facturas duplicadas y consumos positivos del Snapshot.

### Puerta 3 — Preparación de migración

- Backup verificable del libro, backend, WAL y repositorio.
- Golden Master preservado.
- Informe de excepciones con correspondencia de GUID, fechas y facturas.
- Plan de reversión probado.
- Transición de cola antigua sin pérdida ni duplicación.

### Puerta 4 — Implementación autorizada

Solo se autorizarán cambios quirúrgicos después de superar las puertas 1–3. La implementación se hará por fases, con inspección exacta de código antiguo, dependencias y llamadores; compilación en Excel 2016; pruebas dirigidas y confirmación del usuario entre pasos cuando se trate de cambios de producción.

## 15. Pruebas de aceptación

### Órdenes e identidad

1. Introducir A sin B; comprobar GUID y borrador WAL durable.
2. Editar DNI, nombre, fechas y campos funcionales antes de B; comprobar que no se envía UPDATE a Access y que los datos sobreviven al refresco.
3. Introducir B; confirmar reserva atómica por `(Residencia, AnioOrden, NumOrden)`.
4. Introducir A en diciembre y B en enero del ejercicio siguiente; el año procede de B.
5. Dos usuarios reservan la misma clave; solo uno confirma.
6. Simular pérdida de respuesta después del Commit; repetir la misma operación no crea otro expediente.
7. Conflicto de número; aceptar otro número conserva GUID y usa nuevo IdOperacion.
8. Archivar y reutilizar el mismo número en el mismo ejercicio; se genera nuevo GUID y ambos expedientes se conservan.
9. Modificar B a otro ejercicio en orden confirmada; validar nueva clave y rollback completo si hay conflicto.

### Facturas

10. Emitir factura con FechaSalida de ejercicio distinto a FechaPeticion.
11. Dos usuarios solicitan simultáneamente siguiente factura; no existe duplicidad.
12. Repetir la misma operación FIFO tras un Commit no consume otro número.
13. Asignar manualmente un número ya consumido; se rechaza sin alterar Excel.
14. Vaciar la asociación vigente; el consumo permanece en el ledger.
15. Corregir FechaSalida a otro ejercicio con factura: edición ordinaria bloqueada y operación guiada preserva el consumo antiguo.
16. Reemitir factura; nuevo IdOperacion, nueva fila de consumo y anterior conservada.
17. La misma tupla física no se puede vincular a dos GUID distintos.

### Histórico y campos

18. Probar todas las ramas del umbral 17 y los nueve obligatorios.
19. Probar exactamente los dos supuestos de discrepancia C/P y verificar que estados ajenos a los conjuntos no bloquean.
20. Archivar, refrescar y recuperar los cinco campos.
21. Una instancia GUID no sobrescribe el histórico de otra con igual sede/año/número.
22. Vaciados completos limpian AJ; fórmulas protegidas siguen intactas; saneamiento empieza en AK.

### Cola, concurrencia y ordenación

23. Desconectar después de A, capturar B y volver a conectar; un solo expediente confirmado.
24. Reprocesar WAL después de un Commit y antes del marcado local; no duplica operaciones ni facturas.
25. Migrar una fila pendiente del formato antiguo; si es ambigua, queda en cuarentena sin perderse.
26. Refrescar con borradores locales; no desaparecen ni se asocian a otra fila.
27. Ordenar desde AutoFilter y Datos → Ordenar en las tres residencias; A:AJ permanece unido.
28. Lease válido, expirado, de otra instancia y pérdida de lease; ninguna edición se confirma indebidamente.

### Migración

29. Conciliar filas sin FechaPeticion sin inventar ejercicio.
30. Conciliar facturas activas, históricas y snapshot antes de crear unicidad.
31. Verificar el esquema de destino y `ContadorFacturas.Ejercicio = dbLong`.
32. Compilar en Excel 2016 de 32 bits y probar ACE/DAO.

---

## Anexo A — Criterios de no regresión

- No cambiar la disposición visible de las hojas ni el procedimiento de entrada A → B.
- No utilizar `Year(Now)` como sustituto de `FechaPeticion` o `FechaSalida`.
- No usar el número de orden como identidad de fila.
- No sobrescribir un GUID con otro por reintento.
- No descartar WAL pendiente ni borrar operaciones no conciliadas.
- No emitir UPDATE de un expediente que aún no tenga reserva Access confirmada.
- No consumir dos veces la misma factura por reintento.
- No permitir que un número físico de factura se asocie a dos expedientes.
- No eliminar consumos históricos para “liberar” números.
- No volver a crear tablas N:M de facturación: no existen facturas agrupadas.
- No hacer backfill de los cinco campos: el usuario confirma que no existe fuente persistente.
- No ampliar `UsuarioModificacion` ni modificar el token de sesión de 48 bits sin evidencia.
- No ordenar parcialmente A:AB si existen datos en columnas posteriores.
- No limpiar AJ mediante el saneamiento AK+; AJ requiere borrado explícito.
- No utilizar `SELECT *` para transferir registros entre tablas con esquemas evolutivos.

## Anexo B — Fuentes y límites de evidencia

Esta especificación consolida los informes de las Órdenes 1R y 1S–1X y el borrador arquitectónico de 19-11-2026. Los datos de backend citados en auditorías anteriores corresponden al entorno de pruebas/laboratorio, no deben confundirse con una lectura viva de producción.

La comprobación de compatibilidad de validaciones Excel para `TAG`/`R`/`REPARTIR COSTE`, el recuento de facturas positivas y ejercicios de `BDAS_Historico_Snapshot`, y los pendientes actuales de `SYS_COLA_SYNC`/WAL permanecen como una única inspección local de solo lectura previa al laboratorio. No se afirma que se hayan ejecutado en esta sesión porque no se proporcionaron los binarios Access/Excel ni los WAL por estación.

**Conclusión:** la arquitectura funcional y lógica de v1.0 queda definida. La ejecución del proyecto pasa ahora a la inspección local final y, solo después, al laboratorio. Este documento no habilita por sí solo cambios en producción.
