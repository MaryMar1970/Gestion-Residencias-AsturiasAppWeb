# RESIDENCIAAPP
# DOCUMENTO DE DISEÑO TÉCNICO — ARQUITECTURA OBJETIVO DE NUMERACIÓN, FACTURACIÓN Y ARCHIVADO

**Versión:** 0.2 — Borrador arquitectónico revisado tras notas críticas  
**Fecha:** 19 de noviembre de 2026  
**Ámbito:** Excel 2016 (front-end VBA) + Microsoft Access / ACE OLEDB (back-end)  
**Estado:** Diseño revisado; pendiente de aprobación expresa del usuario antes de cualquier implementación.  
**Intervención realizada:** Diseño documental. No se ha modificado código, libro Excel ni base de datos.

---

## 0. Finalidad y límites del documento

Este documento define la arquitectura objetivo para cuatro áreas relacionadas de ResidenciaApp:

1. Identidad y numeración de las solicitudes mediante Nº ORDEN.
2. Numeración independiente de las facturas, cuando proceda.
3. Conservación y sustitución de expedientes en `Ordenes_Historico`.
4. Borrados, archivados, concurrencia y sincronización online/offline.

El diseño parte de las reglas funcionales expresamente confirmadas por el usuario y de las evidencias de las Órdenes 1Q, 1Q-A y 1R.

**No es una autorización de implementación.** No contiene instrucciones para alterar el backend ni el libro activo. Los campos e índices descritos como objetivo son propuestas arquitectónicas pendientes de aprobación y de una comprobación de compatibilidad en una copia de laboratorio.

### 0.1. Principios de gobierno

- ChatGPT ejerce las funciones de arquitecto técnico y auditor: define las reglas, el diseño, los contratos, las transacciones, los errores y las pruebas.
- Antigravity actúa únicamente como inspector/extractor de evidencias o como ejecutor/verificador de instrucciones concretas ya autorizadas.
- Antigravity no decide arquitectura, algoritmos, SQL, código ni reglas funcionales.
- No se reescriben módulos completos ni se reabren decisiones aceptadas sin evidencia nueva y concreta.
- Antes de cualquier futura intervención sobre código o esquema se exige copia de seguridad y verificación de las fuentes exactas.
- La implementación deberá ser quirúrgica, compatible con Excel 2016 y respetar el funcionamiento y disposición actuales de las hojas.

---

# PARTE I — REGLAS FUNCIONALES APROBADAS

## 1. Modelo de identidad de las órdenes

### 1.1. Nº ORDEN

Toda solicitud tiene un Nº ORDEN.

El ejercicio de la orden se obtiene exclusivamente de la fecha de petición:

- Excel: columna **B — FECHA PETICIÓN**.
- Access: campo `FechaPeticion`.
- Regla: `AnioOrden = Year(FechaPeticion)`.

La clave funcional de una orden es:

`(Residencia, AnioOrden, NumOrden)`

Consecuencias:

- La numeración se reinicia por residencia y ejercicio.
- El mismo número numérico puede utilizarse en ejercicios diferentes.
- En `Ordenes` no puede haber dos reservas activas con la misma combinación residencia/año/Nº ORDEN.
- `Ordenes_Historico` no debe bloquear la creación de una nueva reserva activa con esa misma clave.
- Una misma clave puede estar presente simultáneamente una vez en `Ordenes` y una vez en `Ordenes_Historico`: representan la reserva activa nueva y la mejor versión histórica conservada.
- En `Ordenes_Historico` debe existir como máximo una versión por `(Residencia, AnioOrden, NumOrden)`.

La reutilización tras un borrado o archivado significa que el histórico no actúa como restricción de la nueva asignación. El índice único de la tabla activa sí impide que existan dos reservas activas iguales.

### 1.2. Requisito de fecha

Como el ejercicio se deriva de `FechaPeticion`, no se puede confirmar una reserva con Nº ORDEN definitivo si `FechaPeticion` no contiene una fecha válida.

Una fila sin `FechaPeticion` puede existir temporalmente en una fase de edición solo si el flujo funcional lo permite; no puede consolidarse como una orden válida en Access ni recibir una clave anual indeterminada.

El procedimiento de asignación debe validar la fecha antes de reservar el Nº ORDEN. La validación debe repetirse en el servidor Access durante la operación transaccional.

### 1.3. Reutilización y elección del siguiente número

La regla de negocio permite reutilizar un número cuyo expediente anterior ya no esté en `Ordenes`, aunque exista una versión en `Ordenes_Historico`.

Esto no significa necesariamente que el asignador automático vaya a buscar siempre el menor hueco. El código actual dispone de una rutina `ReservarSiguienteOrdenLibreBD` cuya estrategia existente se apoya en el máximo de la tabla activa. La decisión sobre si la asignación automática debe seguir usando `MAX(NumOrden)+1` dentro del ejercicio o buscar el primer hueco libre es distinta de la autorización para introducir manualmente un número ya liberado. No se cambia la estrategia automática sin una decisión expresa.

## 2. Modelo de identidad de las facturas

### 2.1. Serie independiente

Nº ORDEN y Nº FACTURA son dos series independientes, con contadores y ejercicios distintos.

- Toda solicitud tiene Nº ORDEN.
- No todas las solicitudes generan factura.
- Un expediente puede tener `NumFactura` vacío/NULL.
- Solo se consume un número de factura cuando se emite una factura.
- Una factura queda asociada al Nº ORDEN de su expediente en la misma fila de `Ordenes`.
- Cuando la fila se archiva, la relación Nº ORDEN/Nº FACTURA se traslada con el expediente.

### 2.2. Ejercicio de factura

El ejercicio de la factura se obtiene exclusivamente de la fecha de salida:

- Excel: columna **M — SALIDA**.
- Access: campo `FechaSalida`.
- Regla: `AnioFactura = Year(FechaSalida)`.

La fecha del reloj del ordenador (`Year(Now)` o `Year(Date)`) no puede sustituir a `Year(FechaSalida)` para calcular el ejercicio de una factura.

Ejemplo confirmado:

- Solicitud presentada en diciembre de 2027.
- Nº ORDEN 150 del ejercicio 2027.
- Salida en enero de 2028.
- Primera factura emitida del ejercicio de salida: Nº FACTURA 1/2028.

### 2.3. Clave funcional de factura

`(Residencia, AnioFactura, NumFactura)`

La clave incluye residencia y ejercicio. Por ello, el mismo número numérico puede existir en dos ejercicios distintos o en residencias diferentes.

Una misma combinación residencia/año/Nº FACTURA no puede asociarse a dos Nº ORDEN distintos. La hoja activa debe impedir la duplicidad en la columna C dentro del ejercicio de facturación correspondiente; Access debe proporcionar una segunda protección frente a concurrencia. El mecanismo definitivo para mantener esta unicidad al coexistir una versión activa y otra histórica está pendiente del punto 11.

### 2.4. Ausencia de factura

El estado “sin factura” se representa como campo vacío/NULL, no como un número de factura válido.

El valor `0` no es un número válido de factura: la numeración comienza en 1. El informe 1R detectó cuatro valores `NumFactura=0` en la base de pruebas de Oviedo. No se deben modificar esos datos de prueba; el nuevo código debe impedir que se generen nuevos ceros y la eventual normalización de datos existentes deberá tratarse en una migración autorizada.

### 2.5. Contador independiente

El contador de facturas se mantiene por `(Residencia, Ejercicio)`, independiente del contador de órdenes.

La tabla existente `ContadorFacturas` ya tiene una clave única sobre `(Residencia, Ejercicio)` y un campo `UltimoNumero`. Debe corregirse la inicialización de ejercicios nuevos: actualmente, si falta la fila del ejercicio, el `UPDATE` afecta a cero registros y el procedimiento termina fallando.

El contador es un generador secuencial; no es un registro individual de facturas ni guarda por sí solo la relación con un expediente.

---

## 3. Reglas de archivado por residencia

### 3.1. Columnas de datos y fórmulas

El recuento se hace sobre los valores reales de las celdas, excluyendo columnas con fórmulas aunque muestren resultados.

La enumeración extraída de las hojas es la siguiente:

| Residencia | Columnas que cuentan como datos | Columnas excluidas por contener fórmulas | Total de columnas de datos |
|---|---|---|---:|
| GIJÓN | A:T, V:AA, AD | U, AB, AC | 27 |
| SOTO | A:P, S:T, V:AA, AD | Q, R, U, AB, AC | 25 |
| OVIEDO | A:U, W:AE | V, AF, AG | 30 |

**Importante:** estos totales se recalcularon a partir de los intervalos enumerados en el informe 1Q-A. En aquel informe aparecían totales aritméticos inferiores en una unidad. Para diseñar el recuento se utilizarán las listas de columnas, no los totales erróneos del informe.

La clasificación debe verificarse de nuevo contra el código/hojas activas justo antes de implementar, porque las celdas de datos y fórmulas forman parte de la regla funcional.

**Bloqueo técnico adicional acreditado:** `SincronizarFilaHaciaAccess` actualmente lee solo `A:AB` (28 columnas). El mapa 1Q-A identifica columnas de datos fuera de ese rango: `AD` en GIJÓN/SOTO y `AC:AE` en OVIEDO. Por tanto, el backend actual no dispone necesariamente de todos los valores que exige el recuento de archivado. No basta con aumentar ciegamente la lectura hasta `AD`/`AG`: se debe leer el conjunto exacto de columnas de datos de cada sede, excluir fórmulas y verificar qué campos de esas columnas tienen correspondencia en Access. Las columnas fuera de `A:AB` deben incorporarse a la persistencia y a la sincronización inversa/actualizaciones parciales/offline cuando representen datos de usuario que deben conservarse. El recuento y la evaluación de los nueve obligatorios no se implementarán hasta cerrar ese mapa de persistencia.

### 3.2. Nueve campos obligatorios

Además del umbral de cumplimentación, los nueve campos indicados deben estar presentes para que un expediente pueda conservarse en `Ordenes_Historico`.

| Contenido | GIJÓN | SOTO | OVIEDO |
|---|---|---|---|
| Nº ORDEN | A | A | A |
| FECHA PETICIÓN | B | B | B |
| DNI | I | I | I |
| NOMBRE | K | K | K |
| ENTRADA | L | L | L |
| SALIDA | M | M | M |
| RESOLUCIÓN | P | P | P |
| DIRECCIÓN | W | W | X |
| POBLACIÓN | Y | Y | Y |

En particular, DIRECCIÓN es W en GIJÓN y SOTO, pero X en OVIEDO. No se puede reutilizar sin adaptación un único mapa de columnas para las tres sedes.

### 3.3. Qué significa “cumplimentado”

A efectos del diseño:

- Para texto: valor no nulo y no vacío tras eliminar espacios exteriores.
- Para fechas: valor de fecha válido, no nulo.
- Para Nº ORDEN: número válido, no nulo y dentro del rango admitido.
- Para DIRECCIÓN/POBLACIÓN/NOMBRE/DNI: texto no vacío.
- Para RESOLUCIÓN: texto no vacío.
- Las columnas de fórmulas nunca suman al umbral.

La implementación debe utilizar el tipo real de cada columna y evitar contar como cumplimentados valores que sean solo espacios. La comprobación específica del campo A debe validar un Nº ORDEN real.

---

## 4. Coherencia entre Nº FACTURA y RESOLUCIÓN

Se han definido dos discrepancias que impiden considerar correcto el expediente para archivado.

### 4.1. Discrepancia tipo 1

Existe Nº FACTURA en C y la RESOLUCIÓN en P es cualquiera de estos valores:

- `NO`
- `DENEGADA`
- `DESESTIMADA`
- `RENUNCIA`

### 4.2. Discrepancia tipo 2

La celda C está vacía/sin factura y la RESOLUCIÓN en P es cualquiera de estos valores:

- `SI`
- `CONCEDIDA`
- `REEVALUADA`

Las comparaciones deben normalizar espacios exteriores y diferencias entre mayúsculas/minúsculas. Los valores deben compararse con las categorías aprobadas; no se debe inferir equivalencia con otras etiquetas que no hayan sido autorizadas.

### 4.3. Tratamiento de las discrepancias

No se permitirá introducir en el histórico un expediente con una de estas discrepancias cuando la fila haya alcanzado todos los requisitos de admisión (17 o más datos y los nueve obligatorios).

En tal caso:

- Se muestra un aviso explicando el dato concreto que no es coherente.
- No se ofrece una opción de “archivar de todas formas”.
- Se cancela la operación de borrado/archivado.
- La fila permanece intacta en `Ordenes`.
- No se modifica ni elimina una fila anterior de `Ordenes_Historico`.

Mensaje base aprobado:

> **ARCHIVADO CANCELADO.** Debe revisar y resolver la discrepancia entre el Nº FACTURA y la RESOLUCIÓN del expediente antes de volver a intentar su eliminación o archivado.

Se debe añadir una explicación específica: hay Nº FACTURA con resolución negativa, o no hay Nº FACTURA con resolución positiva.

---

## 5. Matriz de decisiones de borrado y archivado

Esta matriz sustituye los criterios anteriores de 15/20 campos, de “datos reales = DNI o Nombre” y de autorización para archivar una discrepancia.

El recuento de nueve obligatorios se interpreta así:

- 0 a 6 presentes: faltan más de dos.
- 7 u 8 presentes: falta uno o dos.
- 9 presentes: están todos.

| Nº de campos de datos | Obligatorios presentes | Discrepancia C/P | Decisión aprobada |
|---:|---:|---|---|
| 0–16 | Cualquiera | Cualquiera | Borrado definitivo, sin aviso |
| 17 o más | 0–6 | Cualquiera | Borrado definitivo, sin aviso |
| 17 o más | 7–8 | Sí | Borrado definitivo |
| 17 o más | 7–8 | No | No archivar; aviso para elegir eliminar definitivamente o modificar |
| 17 o más | 9 | Sí | Cancelar operación; aviso ARCHIVADO CANCELADO; conservar en `Ordenes` |
| 17 o más | 9 | No | Expediente apto para archivado |

### 5.1. Aviso para expedientes con 7 u 8 campos obligatorios y sin discrepancia

El sistema presenta un emergente que permite elegir:

- **Eliminar:** borrar definitivamente de `Ordenes`, sin crear/modificar histórico.
- **Modificar:** cancelar la operación y mantener el expediente activo en `Ordenes`.

No se archiva en ninguno de los dos resultados de este aviso.

### 5.2. Borrado definitivo

El borrado definitivo elimina la fila activa y no crea ningún registro histórico ni rastro de auditoría de esa asignación. El histórico anterior de otra versión del mismo expediente no debe borrarse por este motivo.

### 5.3. Prioridad y atomicidad

La matriz se aplica coherentemente a las vías de borrado individual, lote y toda la residencia. No puede existir una ruta alternativa que siga archivando con el criterio antiguo DNI/Nombre.

La decisión funcional debe calcularse antes de eliminar datos. Si se necesita mostrar un diálogo, no se debe mantener abierta una transacción de base de datos mientras el usuario decide. Tras la respuesta, el procedimiento debe volver a adquirir el bloqueo y leer de nuevo la fila para comprobar que no haya cambiado antes de ejecutar la operación transaccional definitiva.

---

## 6. Selección de la versión que permanece en el histórico

La unidad de comparación es:

`(Residencia, AnioOrden, NumOrden)`

La versión candidata solo es admisible si cumple simultáneamente:

1. Al menos 17 columnas de datos cumplimentadas.
2. Los nueve campos obligatorios de su residencia.
3. Coherencia entre Nº FACTURA y RESOLUCIÓN.

La regla actual de selección aprobada es:

- Si no existe versión histórica, insertar la nueva.
- Si existe una versión histórica y la nueva candidata es admisible, la nueva versión es la última archivada y sustituye a la anterior.
- Si ambas versiones cumplen, prevalece la última en grabarse en `Ordenes_Historico`, independientemente de cuál tenga más campos cumplimentados.
- Si la nueva versión no es admisible, no debe desplazar una versión histórica anterior válida.
- Una nueva fila que se elimina definitivamente no altera el histórico anterior.

**Se elimina expresamente el criterio anterior de “conservar la que tenga mayor número de campos cumplimentados”.** El umbral de 17 es un filtro de admisión, no una puntuación para decidir cuál gana.

La sustitución de una fila histórica válida por una nueva versión válida implica sustituir el contenido del expediente, incluida su asociación Nº FACTURA/FechaSalida. Esta consecuencia debe ser visible en las pruebas: si la versión nueva no tiene factura y la anterior sí, la asociación anterior desaparecerá junto con la versión histórica sustituida. No se debe conservar parcialmente una factura antigua en una fila nueva si ello rompe la coherencia C/P.

---

# PARTE II — ARQUITECTURA TÉCNICA OBJETIVO PROPUESTA

## 7. Situación física actual acreditada

### 7.1. Tabla `Ordenes`

El backend inspeccionado tiene 43 campos. `Id` es autonumérico y clave primaria. Existe un índice único de negocio sobre `(Residencia, NumOrden)` llamado `idx_Residencia_NumOrden`. Existen además índices secundarios por DNI, FechaEntrada, FechaSalida, Residencia y Resolucion.

No existe `AnioOrden` ni `AnioFactura`.

### 7.2. Tabla `Ordenes_Historico`

Tiene los mismos 43 campos, nombres, tipos y orden físico que `Ordenes`. `Id` es autonumérico, pero no tiene clave primaria ni índice único. Su único índice es `idx_Hist_Res_NumOrden` sobre `(Residencia, NumOrden)`, no único.

No existen campos `AnioOrden` ni `AnioFactura`.

### 7.3. Tabla `ContadorFacturas`

Existe con los campos:

- `Id` autonumérico.
- `Residencia` texto corto (20).
- `Ejercicio` entero largo según extracción DAO 1Q-A; el informe también lo identifica en un apartado como entero de 2 bytes, por lo que el tipo exacto debe confirmarse antes de cualquier migración.
- `UltimoNumero` entero largo.

Tiene índice único sobre `(Residencia, Ejercicio)`. En el backend de pruebas de Orden 1Q-A había tres filas, todas del ejercicio 2026.

### 7.4. Ausencia de relaciones físicas

La inspección del catálogo Access comunicada en el Documento Maestro no encontró claves foráneas físicas. No se debe presuponer una relación Access padre-hijo entre `Ordenes` y `Ordenes_Historico`.

### 7.5. Snapshot

`BDAS_Historico_Snapshot` contiene también `NumFactura`, pero no es un registro individual de facturas emitidas. Su función exacta y los procesos de escritura/restauración deben comprobarse antes de decidir si requiere los nuevos campos de ejercicio.

---

## 8. Cambios de esquema propuestos

### 8.1. `Ordenes`

Propuesta de añadir:

| Campo | Tipo propuesto | Regla |
|---|---|---|
| `AnioOrden` | LONG | Obligatorio para toda orden consolidada; debe coincidir con `Year(FechaPeticion)` |
| `AnioFactura` | LONG, admite NULL | Se informa cuando existe factura y debe coincidir con `Year(FechaSalida)` |

Índices propuestos:

| Nombre lógico | Campos | Propiedad |
|---|---|---|
| `UX_Ordenes_Orden` | `Residencia, AnioOrden, NumOrden` | Único; los tres campos deben estar informados en toda orden consolidada |
| Índice de consulta de factura | `Residencia, AnioFactura, NumFactura` | No único; facilita búsquedas y validaciones, pero no es la barrera global de unicidad |

No se propone un índice único compuesto de factura sobre esta tabla porque `NumFactura` y `AnioFactura` son NULL cuando no hay factura, mientras `Residencia` siempre tiene valor. La propiedad DAO `IgnoreNulls` no debe utilizarse como si fuera un índice filtrado por `NumFactura IS NOT NULL`; el comportamiento de claves compuestas con NULL no debe convertirse en una dependencia del diseño. La unicidad global se controlará mediante `RegistroNumerosFactura` (apartado 8.5).

El índice único actual `idx_Residencia_NumOrden` debe retirarse o sustituirse en la migración, porque impediría que el mismo Nº ORDEN se utilice en distintos ejercicios.

`Id` continuará como clave primaria técnica de `Ordenes`.

### 8.2. `Ordenes_Historico`

Propuesta de añadir los mismos campos de ejercicio:

- `AnioOrden`.
- `AnioFactura`.

Índices propuestos:

| Nombre lógico | Campos | Propiedad |
|---|---|---|
| `UX_Historico_Orden` | `Residencia, AnioOrden, NumOrden` | Único: como máximo una versión por clave de orden histórica |
| Índice de consulta de factura | `Residencia, AnioFactura, NumFactura` | No único; búsqueda y conciliación |

No se crea un índice único compuesto sobre los campos opcionales de factura en esta tabla. La unicidad de la factura se garantiza mediante la tabla común `RegistroNumerosFactura`, que contiene exclusivamente asignaciones reales y cuyos campos de clave son obligatorios.

El índice histórico de orden permite una versión por clave de orden dentro del histórico, pero **no impide la misma clave en `Ordenes`**, porque los índices se aplican independientemente a cada tabla.

Se recomienda dotar a `Ordenes_Historico.Id` de una clave primaria propia si las pruebas de migración confirman que no hay valores repetidos/nulos y que no existe dependencia VBA que necesite la ausencia actual de PK. Esta modificación no se considera imprescindible para definir la clave funcional; queda sujeta a análisis de impacto.

No se propone copiar ciegamente con `SELECT *`. La transferencia deberá enumerar explícitamente campos de origen y destino para evitar errores al ampliar las tablas.

### 8.3. `ContadorFacturas`

Conservar la tabla y el índice único `(Residencia, Ejercicio)`.

Cambios funcionales necesarios:

1. Inicialización segura del contador para un ejercicio sin fila previa.
2. Asignación atómica del siguiente número, con actualización transaccional.
3. Sincronización del contador después de una asignación manual.
4. No decrementar `UltimoNumero` por borrados, sustituciones históricas o purgas.
5. No usar `MAX(NumFactura)` sin filtrar ejercicio ni como sustituto de la reserva atómica.
6. El ejercicio recibido por todas las vías debe venir de `Year(FechaSalida)` del expediente y nunca de `Year(Now)`.

No se propone crear una tabla de contadores de órdenes en esta fase. Se aprovechará la reserva transaccional existente, adaptándola a `AnioOrden`.

### 8.4. Reglas de NULL para facturas

- No factura: `NumFactura = NULL` y `AnioFactura = NULL`; no se crea fila en `RegistroNumerosFactura`.
- Factura asignada: `NumFactura` entero positivo y `AnioFactura = Year(FechaSalida)`; se crea su registro de control en la misma transacción.
- `NumFactura = 0` no representa “sin factura”; la nueva aplicación debe evitar generarlo.
- Una factura positiva sin `FechaSalida` válida no se puede consolidar.

No dependeremos de `IgnoreNulls` en un índice compuesto que incluya `Residencia`, porque ese campo siempre tiene valor. El diseño evita el problema al mantener la clave única en una tabla separada que solo contiene facturas asignadas, con los tres campos obligatorios y no nulos.

### 8.5. Tabla auxiliar técnica `RegistroNumerosFactura` — registro de unicidad global

Se propone incorporar una tabla técnica de control, no una nueva entidad de facturación ni una tabla de líneas o importes. La fila del expediente seguirá siendo el lugar visible y operativo donde se asocian `NumOrden` y `NumFactura`. La tabla auxiliar existirá únicamente para mantener la unicidad global de los números de factura y conservar constancia de que una combinación ya fue asignada.

Campos mínimos propuestos:

| Campo | Tipo propuesto | Propiedad / finalidad |
|---|---|---|
| `Id` | AutoNumber / LONG | Clave primaria técnica |
| `IdOrden` | LONG | Identificador técnico de la instancia concreta de la orden; obligatorio |
| `Residencia` | Texto (20) | Obligatorio |
| `AnioOrden` | LONG | Obligatorio; año de `FechaPeticion` de esa instancia |
| `NumOrden` | LONG | Obligatorio |
| `AnioFactura` | LONG | Obligatorio; año de `FechaSalida` al asignar la factura |
| `NumFactura` | LONG | Obligatorio y mayor que cero |

Índices propuestos:

- `UX_RNF_ClaveFactura`: único sobre `(Residencia, AnioFactura, NumFactura)`. Este índice impide que la misma factura del mismo ejercicio y residencia se asocie a dos instancias de orden distintas, con independencia de que la orden esté activa o archivada.
- `UX_RNF_IdOrden`: único sobre `IdOrden`, para impedir que una misma instancia de expediente tenga dos facturas simultáneas.

La tabla solo recibe filas cuando se asigna una factura. Por tanto, no contiene registros sin factura ni claves con NULL y no necesita `IgnoreNulls`.

`IdOrden` identifica la instancia concreta, no solo la clave de negocio. Esto es necesario porque el mismo `(Residencia, AnioOrden, NumOrden)` puede coexistir en `Ordenes` y `Ordenes_Historico` como dos solicitudes distintas tras la reutilización permitida del Nº ORDEN. El `Id` técnico que se copia al archivar permite seguir la misma instancia al trasladarla. No se establecerá una clave externa convencional hacia `Ordenes` o `Ordenes_Historico`, porque la instancia puede residir en cualquiera de las dos tablas; la coherencia se mantendrá mediante operaciones transaccionales.

La asociación principal continuará en la fila de `Ordenes`/`Ordenes_Historico`. `RegistroNumerosFactura` será un registro técnico de asignaciones ya consumidas. Si la mejor versión histórica sustituye una versión anterior que tenía factura, el registro técnico de la factura anterior no se elimina: así el número no queda disponible para una futura asignación manual en el mismo ejercicio.

La regla resultante es inequívoca: se permite el mismo número numérico en otro ejercicio o residencia; no se permite volver a asignar la misma combinación `(Residencia, AnioFactura, NumFactura)` dentro del mismo ejercicio, aunque el expediente original haya sido archivado, sustituido o eliminado de las tablas operativas.

Esta tabla es una propuesta estructural nueva y requiere aprobación antes de implementarse. Su persistencia y tratamiento de eventuales rectificaciones/anulaciones de factura deberán mantenerse coherentes con las reglas de negocio que se aprueben; una anulación no debe liberar automáticamente el número para reutilización.

---

## 9. Integridad de Nº ORDEN

### 9.1. Reserva en línea

Se reutilizan las operaciones actuales, adaptándolas a la clave anual:

- `ExisteOrdenEnBD`.
- `IntentarReservarOrdenBD`.
- `ReservarSiguienteOrdenLibreBD`.
- `AsegurarOrdenExiste`.
- `RenumerarOrdenExistenteBD`.
- `SincronizarFilaHaciaAccess`.

El nuevo índice único de `Ordenes` es la protección final frente a dos usuarios que intenten reservar la misma combinación al mismo tiempo. La comprobación previa de VBA sirve para mostrar un aviso comprensible, pero no sustituye al índice.

Toda inserción/renumeración debe calcular `AnioOrden` desde `FechaPeticion` y persistirlo en la misma transacción que `NumOrden` y `Residencia`.

### 9.2. Reutilización

Antes de reservar un número manualmente, se comprueba la clave en `Ordenes` para esa residencia y ejercicio. No se consulta `Ordenes_Historico` como condición de rechazo.

El histórico no debe ser borrado simplemente porque el mismo número se vuelva a utilizar en activo. Puede coexistir con la nueva reserva activa.

### 9.3. Cambio de FechaPeticion

Si cambia `FechaPeticion`, debe recalcularse `AnioOrden`.

La actualización debe comprobar en una sola operación la nueva clave `(Residencia, AnioOrden, NumOrden)`. Si colisiona con otra reserva activa, se rechaza el cambio y se conserva el registro anterior íntegro. La política exacta de renumeración automática ante cambio de ejercicio deberá alinearse con la interfaz actual; no se debe cambiar silenciosamente el Nº ORDEN.

---

## 10. Integridad de Nº FACTURA

### 10.1. Asignación automática

`AsignarFacturaPorColumna` y `ObtenerSiguienteNumFactura` deben dejar de usar `Year(Now)` y obtener el ejercicio desde `FechaSalida`.

El proceso debe:

1. Validar `FechaSalida`.
2. Derivar `AnioFactura`.
3. Adquirir el contador de `(Residencia, AnioFactura)` transaccionalmente.
4. Crear de forma segura la fila del contador si aún no existe, iniciándola de modo que la primera factura sea 1.
5. Incrementar el contador y asociar el nuevo número al expediente concreto.
6. Confirmar conjuntamente el contador y la asignación de la factura en Access.

La transacción debe evitar que dos usuarios obtengan el mismo número. Si la actualización del expediente falla, el incremento de contador no debe quedar confirmado de forma aislada.

### 10.2. Asignación manual

`ProcesarCambioManualFactura`, `SincronizarContadorFacturaManual` y `ValidarFacturaManual` deben aplicar el mismo ejercicio `Year(FechaSalida)` y la misma unicidad.

La asignación manual no puede saltarse el índice único ni el control de correspondencia factura/expediente. Los números no positivos se rechazan.

La regla de tratamiento de un número manual menor o igual que `UltimoNumero` requiere definición final: el contador es un máximo secuencial, pero no es un libro individual de facturas y no demuestra por sí solo si un número inferior fue usado, anulado o quedó libre.

### 10.3. Actualización de FechaSalida después de asignar factura

Si cambia `FechaSalida`, cambia potencialmente `AnioFactura`. La operación debe calcular el nuevo ejercicio y comprobar la clave antes de confirmar la modificación.

No debe quedar una factura con `AnioFactura` distinto de `Year(FechaSalida)`. Si el cambio produce una colisión, la operación debe rechazarse o requerir reasignar la factura; no se debe aceptar una relación inconsistente.

### 10.4. Relación factura/orden

La fila del expediente contiene la relación entre `NumOrden` y `NumFactura`; no se ha aprobado una entidad independiente de factura ni una tabla de líneas de factura.

La asociación se traslada a `Ordenes_Historico` junto con el expediente. No se debe eliminar una factura histórica como efecto colateral de borrar una versión, salvo que una futura regla expresa lo autorice.

---

## 11. Decisión arquitectónica recomendada: unicidad global de facturas

La revisión de las notas críticas permite cerrar la interpretación operativa del punto 11:

1. La clave única de factura es `(Residencia, AnioFactura, NumFactura)`.
2. La combinación completa no puede asociarse a dos instancias distintas de orden, ni aunque una esté en `Ordenes` y otra en `Ordenes_Historico`.
3. Repetir el mismo número numérico en otro ejercicio o residencia sí está permitido porque la clave completa es diferente.
4. Una factura archivada conserva su asociación y su número queda consumido en ese ejercicio.
5. El histórico no bloquea el comienzo de una nueva serie en el siguiente ejercicio; sí impide reemitir la misma clave de factura en el mismo ejercicio.
6. La asignación manual se somete a la misma unicidad global que la automática.

### 11.1. Mecanismo seleccionado

Se recomienda `RegistroNumerosFactura` (apartado 8.5) como registro técnico común de claves de factura utilizadas. Su índice único físico sobre `(Residencia, AnioFactura, NumFactura)` aplica a toda asignación con independencia de la ubicación del expediente.

Cada alta de factura y su actualización en la fila activa deberán realizarse dentro de la misma transacción Access. El archivado de una orden no crea una segunda asignación de factura: conserva el mismo `IdOrden` y la misma asociación ya registrada. La sustitución de una versión histórica no elimina el registro de consumo de la factura anterior.

El contador `ContadorFacturas` seguirá siendo el generador secuencial por residencia y ejercicio. `RegistroNumerosFactura` será la barrera de unicidad y memoria de asignaciones; no se sustituye una función por la otra.

### 11.2. Tratamiento de la numeración manual

Una factura manual solo se acepta si su clave completa no existe en `RegistroNumerosFactura`. Si el número es mayor que `UltimoNumero`, el contador se eleva a ese valor en la misma transacción. Si es inferior o igual al contador pero nunca se ha registrado, podrá utilizarse como hueco manual disponible; la tabla auxiliar será quien determine si ya fue utilizado. Una factura ya registrada no se puede asignar de nuevo a otra instancia de orden.

La operación debe distinguir una repetición idempotente del mismo valor para el mismo `IdOrden` de una nueva asignación a otra instancia. No se permitirá modificar silenciosamente la factura asociada sin actualizar el registro técnico mediante el procedimiento específico autorizado.

### 11.3. Por qué no se utiliza un índice compuesto nullable con `IgnoreNulls`

`Residencia` siempre está informada, incluso cuando no existe factura. Por tanto, no se utilizará un índice único `(Residencia, AnioFactura, NumFactura)` con `IgnoreNulls=True` como mecanismo para omitir expedientes sin factura. La semántica de `IgnoreNulls` en índices compuestos no equivale a un índice filtrado por `NumFactura IS NOT NULL`; cambiar el orden de los campos no corrige la condición, porque `Residencia` continúa siendo no nula.

La tabla técnica evita esta dificultad: solo contiene facturas asignadas y todos sus campos de clave son obligatorios. No hay filas sin factura que deban excluirse del índice.

La documentación DAO de Microsoft define `IgnoreNulls` como la propiedad que controla la presencia de entradas de índice para registros que contienen valores NULL en los campos indexados; la documentación de índices compuestos confirma que la unicidad se aplica a la combinación de todos los campos. El laboratorio deberá confirmar el comportamiento específico del proveedor ACE instalado, pero la arquitectura no dependerá de esa semántica porque la tabla auxiliar no almacena claves nulas.

### 11.4. Alcance de esta decisión

La propuesta evita que una factura repetida llegue a existir en el conjunto de expedientes activos e históricos por una asignación nueva. También conserva la ocupación del número si una versión histórica que contenía la factura es posteriormente reemplazada por otra versión válida.

La incorporación de `RegistroNumerosFactura` es un cambio de esquema adicional. No se debe implementar hasta que se apruebe esta versión del diseño y se pruebe la migración en una copia aislada.
---

## 12. Integridad y sustitución del histórico

### 12.1. Clave del histórico

Clave funcional única dentro de `Ordenes_Historico`:

`(Residencia, AnioOrden, NumOrden)`

Debe conservar como máximo una versión válida de cada expediente.

### 12.2. Operación transaccional

La operación de archivado deberá ejecutarse en una transacción de Access que incluya:

1. Verificación de que la orden activa existe y que el usuario tiene el bloqueo/lease válido.
2. Lectura de todos los campos necesarios para evaluar el recuento, obligatorios y coherencia C/P.
3. Determinación de la acción conforme a la matriz de la sección 5.
4. Si es archivado: comprobar la versión histórica de la misma clave; sustituirla por la nueva versión admisible, o insertarla si no existe.
5. Verificar filas afectadas y consistencia de datos.
6. Eliminar la fila activa solo cuando la escritura/sustitución histórica haya terminado correctamente.
7. Confirmar la transacción.

Si falla cualquier paso de archivado, ejecutar `Rollback` y conservar la fila activa y el histórico previo sin cambios.

### 12.3. Borrado definitivo

En las categorías cuyo resultado es borrar definitivamente, eliminar únicamente la fila activa. No borrar antes la versión histórica existente.

La operación debe ser idempotente de acuerdo con los contratos actuales, pero no debe interpretar un fallo de bloqueo o conexión como permiso para borrar localmente datos que no se han confirmado en Access.

### 12.4. SQL explícito

Las transferencias futuras deberán nombrar las columnas de origen y destino de forma explícita. Se evitará `INSERT INTO Ordenes_Historico SELECT * FROM Ordenes`, porque la ampliación de las tablas rompería la suposición del mismo número y orden físico de campos.

---

## 13. Integridad multiusuario y concurrencia

### 13.1. Índices como última barrera

Las comprobaciones VBA no son suficientes ante dos clientes concurrentes. Las claves únicas físicas en Access deben ser la última barrera para:

- `(Residencia, AnioOrden, NumOrden)` en `Ordenes`.
- `(Residencia, AnioOrden, NumOrden)` en `Ordenes_Historico`.
- `(Residencia, AnioFactura, NumFactura)` en `RegistroNumerosFactura`, con todos sus campos obligatorios y un índice único global de asignaciones.

### 13.2. Reutilización de leases existentes

Se aprovechará el sistema de leases de V5.2, pero antes de autorizar la implementación deberán corregirse o justificarse las debilidades concretas ya detectadas:

- La condición `DateDiff('s', UltimaModificacion, Now()) < 0` puede permitir tomar un bloqueo por diferencia de reloj.
- La alternativa por nombre de usuario de Windows puede permitir que otra instancia del mismo usuario tome el bloqueo de la primera.
- El token creado con `CoCreateGuid` conserva actualmente solo 48 bits hexadecimales del GUID; no debe describirse como UUID completo de 128 bits.

No se debe reescribir todo el sistema de bloqueos. Solo se tocará el contrato de lease en las rutas que intervengan en asignación, modificación, borrado, archivado y sincronización.

### 13.3. Diálogos y transacciones

No mantener una transacción DAO/ADO abierta mientras se espera una respuesta del usuario en un emergente.

Para el caso de elegir eliminar o modificar:

1. Leer la fila y determinar la rama de la matriz sin mantener abierta una transacción mientras se presenta el emergente.
2. Adquirir o renovar el lease con la implementación realmente presente en el libro activo; confirmar el bloqueo y cerrar cualquier transacción de adquisición antes de mostrar el emergente.
3. Mostrar el diálogo y recoger la decisión del usuario fuera de una transacción Access.
4. Tras la respuesta, volver a validar/adquirir el lease de la misma instancia y releer la fila desde Access.
5. Comparar la versión leída con la que originó el diálogo y recalcular recuento, obligatorios y coherencia C/P.
6. Ejecutar la operación transaccional final solo si el bloqueo sigue perteneciendo a esa instancia y la fila no cambió de forma material.

El lease actual tiene una duración nominal de 120 segundos; el diálogo puede permanecer abierto más tiempo. Por ello, no se debe suponer que el bloqueo sigue vigente tras cerrar el emergente: debe renovarse/revalidarse. Si ha caducado y otra instancia lo ha adquirido, se cancela la operación y se refresca la fila.

No se debe depender de `clsDatabaseEngine` como componente de producción sin acreditar su presencia y homologación en el binario activo. La base de trabajo será la implementación V5.2 efectivamente presente en el proyecto activo. Antes de usarla como protección del diálogo deben corregirse/verificarse las debilidades ya detectadas: apropiación por `DateDiff < 0`, alternativa de propiedad por nombre de usuario de Windows entre instancias del mismo usuario y truncamiento del GUID a 48 bits.

---

## 14. Sincronización offline

El procedimiento real inspeccionado es `SincronizarColaOfflineHaciaAccess` en `ModuloSeguridadLocal.bas`, con la cola local `SYS_COLA_SYNC` y persistencia WAL.

Reglas objetivo:

- Los números provisionales de orden se mantienen dentro de la estrategia existente (rango desde 900001) mientras no haya conexión.
- Al recuperar red, la resolución del provisional debe determinar `AnioOrden` desde la `FechaPeticion` original de la solicitud, no desde la hora de sincronización.
- La conciliación de facturas debe derivar `AnioFactura` desde `FechaSalida`, nunca de la marca temporal del WAL ni de `Year(Now)`.
- La reconciliación debe ser idempotente: reintentar un evento no puede crear otra orden ni asociar otra factura.
- Los eventos diferidos deben conservar la relación con la clave de expediente correcta cuando se resuelva un número provisional.
- Si la operación online detecta duplicidad, fecha inválida o conflicto de factura, el evento debe quedar sin confirmar y seguir el tratamiento de error definido; no se debe marcar como sincronizado falsamente.

No se debe cambiar el formato WAL ni el orden de columnas sin un análisis concreto de compatibilidad y una orden expresa.

---

## 15. Compatibilidad con `BDAS_Historico_Snapshot` y purga RGPD

`BDAS_Historico_Snapshot` contiene `NumFactura`, pero no constituye un libro de facturas. Antes de implementar, hay que verificar qué procedimientos copian, restauran o consultan el snapshot.

La decisión de incorporar `AnioOrden` y `AnioFactura` al snapshot dependerá de si participa en la restauración de expedientes o solo conserva una fotografía de soporte. No se deben añadir campos por mera simetría sin comprobar dependencias.

`PurgarHistoricoRGPD` es una operación distinta del archivado ordinario: elimina físicamente registros históricos conforme a su política de retención. No debe mezclarse con el algoritmo de selección de la mejor versión histórica. Cualquier purga debe respetar la política legal y no realizar liberaciones accidentales de contadores de facturas.

---

# PARTE III — PLAN DE CAMBIOS Y VALIDACIÓN

## 16. Procedimientos existentes que probablemente requieren adaptación

La siguiente lista es un mapa inicial basado en la Orden 1Q; no es una autorización para modificarlos. Antes de escribir el cambio se comprobará la firma y el código exacto en la fuente activa.

### 16.1. Base de datos y órdenes

- `ExisteOrdenEnBD`.
- `IntentarReservarOrdenBD`.
- `ReservarSiguienteOrdenLibreBD`.
- `AsegurarOrdenExiste`.
- `RenumerarOrdenExistenteBD`.
- `SincronizarFilaHaciaAccess`.
- `ActualizarOrden`.
- `ActualizarOrdenMultiple`.
- `EliminarOrdenBD`.
- `EliminarOrdenesLoteBD`.
- `EliminarTodasOrdenesSedeBD`.
- `ArchivarOrdenesAntiguasBD`.
- `SincronizarHojaDesdeAccess`.

### 16.2. Facturación

- `AsignarFacturaPorColumna`.
- `ProcesarCambioManualFactura`.
- `ValidarFacturaManual`.
- `ObtenerSiguienteNumFactura`.
- `SincronizarContadorFacturaManual`.
- `ExisteFacturaEnBD`.
- `ConsolidarFacturaDiferidaFIFO`.

### 16.3. Offline y borrados

- `SincronizarColaOfflineHaciaAccess`.
- `ResolverNumOrdenProvisionalEnBD`.
- Encolado y consolidación de operaciones de borrado individual y en lote.
- Llamadores de las tres rutas de borrado de la aplicación.

No se debe asumir que esta lista recoge todas las dependencias indirectas. La implementación requiere una extracción dirigida de los llamadores, variables y SQL afectados solo cuando el diseño esté aprobado.

## 17. Secuencia de implementación propuesta

La implementación se dividirá en pasos pequeños, con copia de seguridad, verificación y prueba antes de avanzar.

| Paso | Alcance | Puerta de aceptación |
|---|---|---|
| 0 | Congelar fuentes, copia de seguridad y comprobar hashes | Fuentes inequívocas y respaldo verificable |
| 1 | Migración de esquema en copia de laboratorio: campos e índices | Esquema consistente; sin pérdida de campos ni registros de prueba |
| 2 | Reserva y validación de Nº ORDEN por ejercicio | Unicidad bajo concurrencia y reutilización permitida tras salir de activa |
| 3 | Contador y asignación de facturas por FechaSalida | Numeración anual independiente y sin duplicidades según regla finalmente aprobada |
| 4 | Matriz centralizada de borrado/archivado | Todos los casos de la sección 5 reproducibles |
| 5 | Sustitución histórica y transferencia de asociación de factura | El histórico contiene solo la última versión válida por clave |
| 6 | Integración offline y leases | Reintentos idempotentes; sin duplicados ni toma indebida de bloqueos |
| 7 | Pruebas de regresión dirigidas en Excel 2016 | Flujos actuales conservados y cero errores COM/VBA |

No pasar al siguiente paso sin aceptación del resultado del paso anterior por el usuario.

## 18. Pruebas mínimas de aceptación

### 18.1. Nº ORDEN

1. Crear GIJÓN/2027/101 y volver a intentar la misma clave en otra fila activa: debe rechazarse.
2. Crear GIJÓN/2028/101: debe permitirse.
3. Archivar GIJÓN/2027/101 y crear una nueva activa GIJÓN/2027/101: debe permitirse, sin borrar automáticamente la versión histórica.
4. Repetir la asignación simultáneamente desde dos instancias: solo una debe confirmarse.
5. Cambiar FechaPeticion a otro año con colisión activa: rechazar sin alterar el registro.

### 18.2. Facturas

1. Diez órdenes sin factura y una orden con factura 1: válido.
2. Solicitud de diciembre 2027 y salida enero 2028: orden 2027 y factura 1/2028.
3. Repetir número de factura en otra residencia: permitido.
4. Repetir número en misma residencia y ejercicio, asociado a otro expediente: rechazado.
5. Repetir el mismo número en ejercicio distinto: permitido.
6. Introducir `NumFactura=0`: rechazado como valor inválido.
7. Asignar factura con FechaSalida vacía: rechazado.
8. Dos usuarios solicitan simultáneamente la siguiente factura: se asignan números distintos.
9. Reintentar un evento offline: no crea una segunda factura ni modifica la asociación ya confirmada.

### 18.3. Archivado — matriz

Probar, al menos, los seis supuestos de la sección 5, incluyendo:

- 16 campos de datos.
- 17 campos y 6 obligatorios.
- 17 campos, 7 obligatorios, sin discrepancia.
- 17 campos, 7 obligatorios, con discrepancia.
- 17 campos, 9 obligatorios, con discrepancia.
- 17 campos, 9 obligatorios, sin discrepancia.

Verificar para cada uno el estado final de ambas tablas, mensaje emergente, valores de Excel y número de filas afectadas.

### 18.4. Selección histórica

1. No existe histórico: insertar la nueva versión válida.
2. Existe histórico válido y nueva versión válida: gana la última, aunque tenga menos campos que la anterior, siempre que supere el umbral y cumpla todos los obligatorios.
3. Nueva versión inválida: no debe desplazar una versión histórica válida.
4. La operación falla al insertar/sustituir: rollback; fila activa e histórico previo permanecen intactos.
5. La misma clave existe simultáneamente en activo e histórico: ambas pueden coexistir.

### 18.5. Concurrencia y fallos

- Dos usuarios reservando la misma orden.
- Dos usuarios asignando la misma factura manual.
- Bloqueo vigente de otra instancia.
- Mismo usuario Windows desde dos instancias.
- Diferencia de reloj entre equipos.
- Caída de conexión antes y después del `CommitTrans`.
- Reintento de la cola offline.
- Fallo entre inserción histórica y borrado activo.

---

# PARTE IV — PUNTOS PENDIENTES ANTES DE CERRAR EL DISEÑO

## 19. Verificaciones restantes antes de cerrar el diseño

La unicidad global de facturas y la política de asignación manual quedan resueltas en el apartado 11. Los puntos que siguen sí requieren evidencia o una decisión funcional adicional.

### 19.1. Filas antiguas sin fecha

En la base de pruebas de Orden 1R:

- `Ordenes`: 1 fila sin `FechaPeticion`.
- `Ordenes_Historico`: 12 filas sin `FechaPeticion`.

No se puede derivar `AnioOrden` de una fecha inexistente sin inventar el año. Como el backend examinado es de pruebas, no se deben sanear esos registros. Antes de una migración real se deberá inspeccionar la base de datos productiva autorizada y definir el tratamiento de cualquier fila sin fecha.

### 19.2. FechaSalida modificada tras asignar factura

La regla aprobada hace que el ejercicio dependa de `FechaSalida`. Falta precisar el tratamiento de negocio si la fecha se corrige a otro año después de haberse asignado/emitido la factura. La integridad técnica exige no dejar `AnioFactura` desincronizado; la acción de usuario (recalcular, reasignar o bloquear) deberá aprobarse antes de implementar.

### 19.3. Columnas de datos fuera del rango persistido A:AB

El informe 1Q-A clasifica `AD` como dato introducido por usuario en GIJÓN/SOTO y `AC:AE` como datos de usuario en OVIEDO. La evidencia de código disponible indica que `SincronizarFilaHaciaAccess` solo lee A:AB. Debe realizarse una extracción dirigida para identificar los encabezados exactos, campos Access de destino (si existen) y otras rutinas que los persistan o restauren. No se autoriza resolverlo ampliando un rango rectangular sin mapear qué columnas son datos, cuáles fórmulas y cuáles técnicas. Este punto es una condición de entrada para diseñar el recuento de 17 campos y el archivado fiable.

### 19.4. Snapshot

Falta verificar las dependencias de `BDAS_Historico_Snapshot` para decidir si debe recibir los campos `AnioOrden` y `AnioFactura`.

### 19.5. Tipo exacto DAO de `ContadorFacturas.Ejercicio`

El informe 1Q-A presenta una discrepancia en la descripción del ancho/tipo de `Ejercicio` (entero largo frente a entero de 2 bytes en secciones distintas). Debe verificarse en una extracción directa del esquema de la copia de laboratorio que se vaya a modificar. No ejecutar cambios basándose en la descripción ambigua.

### 19.6. Aviso modal para 7 u 8 campos obligatorios

La regla funcional está cerrada (el usuario elige eliminar definitivamente o modificar), pero el texto exacto del emergente y la rotulación de botones se definirán durante la implementación, conservando el significado aprobado.
---

## 20. Evidencias de referencia

Este diseño utiliza los hallazgos contenidos en los informes entregados por Antigravity en esta conversación:

- **ORDEN 1Q:** extracción de código, dependencias y esquema de tablas de órdenes e histórico. Hash del informe comunicado: `97e309398cb1c1e0a23acf8ed62fb24c95714593a5c5aea14beab5435bf58d1e`.
- **ORDEN 1Q-A:** comprobaciones sobre `ContadorFacturas`, `Id`, fechas y columnas de datos/fórmulas. Hash comunicado: `b931e08e7ddfb506da6592c430ff8a1da1c8051e045e26fcd4a248ebe9e956e8`.
- **ORDEN 1R:** controles de fechas, persistencia de facturas y sustitución histórica. Hash del informe comunicado: `5b2393f38a01d18b7e731819d900323ea255fa14d721c69429422264a546d2f8`.

### 20.1. Fuentes y hashes de referencia de Orden 1Q

- `BDAS_Multiusuario.xlsm`: SHA-256 `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`.
- `modDatabase.bas`: SHA-256 `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47`.
- Backend de pruebas en Orden 1Q: SHA-256 `853eb4e736c02934668f0cc10d5a4d393f3cd66ceb40bf887b8993115f0ed430`.

### 20.2. Discrepancia de hash del backend en Orden 1R

El backend de pruebas mostró en Orden 1R SHA-256 `9be5283652ef9f7ae5c4cc85c509f950afc6e90d32350eaf331313aa3fcd2765`, distinto del hash de Orden 1Q. Antigravity atribuyó la diferencia a seis registros ordinarios LOGIN/LOGOUT en `LogActividad` y comunicó que los recuentos de las tablas de negocio eran iguales. No se debe describir el fichero de 1R como binariamente idéntico al de 1Q.

Los datos de estas copias son evidencias de laboratorio/pruebas; no se deben corregir, sanear ni tratar como datos operativos productivos.

---

# 21. Dictamen arquitectónico y siguiente paso

La arquitectura objetivo puede apoyarse en las estructuras existentes y no justifica una reescritura general de ResidenciaApp.

Las modificaciones estructurales principales propuestas son:

1. Guardar explícitamente `AnioOrden` y `AnioFactura` para indexar ejercicios derivados de fechas.
2. Cambiar la clave única de orden activa de `(Residencia, NumOrden)` a `(Residencia, AnioOrden, NumOrden)`.
3. Añadir una clave histórica por `(Residencia, AnioOrden, NumOrden)` para mantener una sola versión histórica por expediente, permitiendo simultáneamente una versión activa con esa misma clave.
4. Incorporar protección de unicidad para la factura, respetando NULL cuando no exista factura y cerrando primero el alcance de la unicidad entre tablas.
5. Corregir la inicialización y actualización transaccional de `ContadorFacturas`.
6. Sustituir las tres lógicas actuales de borrado/archivado por un único contrato funcional que aplique la matriz aprobada.
7. Adaptar la conciliación offline y los bloqueos multiusuario sin reemplazar innecesariamente los mecanismos existentes.

## Próxima actuación recomendada

Antes de redactar código o modificar el esquema, aprobar la tabla auxiliar `RegistroNumerosFactura` y realizar extracciones dirigidas sobre las columnas de datos fuera de A:AB, las dependencias de `BDAS_Historico_Snapshot` y el tipo físico de `ContadorFacturas.Ejercicio`, si esas evidencias no están completas en el repositorio de documentación.

Después se podrá cerrar la versión 1.0 del diseño, emitir la orden de implementación por pasos y exigir una copia de seguridad verificable antes de cualquier cambio.

**Fin del documento.**
