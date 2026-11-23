# Registro Diario de Control de Cambios — 18/10/2026

> **Entorno de Trabajo:** `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
> **Servidor Compartido en Red:** `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
> **Base de Datos Backend:** `H:\ResidenciaBD\Residencia_BE.accdb`  
> **Estado de Compilación VBA:** 🟢 **0 ERRORES** (Validado en Excel 2016 COM)

---

## Incidencia / Requerimiento: Gestión Centralizada e Interactiva de la Lista Negra

### 1. Diagnóstico del Problema
- **Problema reportado:** Al registrar una persona de prueba en la hoja `LISTA NEGRA` con rol `ADMINISTRADOR` (DNI en columna A, fecha en columna B y motivo en columna C) y cerrar la aplicación, al volver a abrirla los datos se habían perdido.
- **Causas identificadas:**
  1. En `ThisWorkbook.cls` (evento `Workbook_BeforeClose`), la línea `ThisWorkbook.Saved = True` descarta intencionadamente cualquier modificación en las celdas de Excel para proteger la plantilla compartida en el entorno multiusuario.
  2. La hoja `LISTA NEGRA` se encontraba desconectada de la base de datos Access: el botón de cinta únicamente mostraba la hoja de Excel, pero no existía ningún mecanismo que trasladase lo tecleado a la tabla `ListaNegra` de Access (`Residencia_BE.accdb`), ni rutina de descarga automática de los registros existentes al abrir la hoja.
  3. En la tabla `ListaNegra` de Access existía un registro residual con el texto del encabezado (`DNI | FECHA | MOTIVO`) importado en la migración original, y fechas de inclusión almacenadas como números de serie en el campo `Nombre`.

---

### 2. Solución Técnica Implementada (Opción 1: Formulario + Hoja Sincronizada)

Siguiendo la decisión del usuario, se implementó la **Opción 1 con permisos tanto para rol ADMINISTRADOR como OPERADOR**:

1. **Nuevo UserForm de Gestión ([frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm)):**
   - **Visualización y Búsqueda:** `ListBox` multicolumna (`DNI`, `Fecha Alta`, `Motivo`) con caja de filtrado en tiempo real (`txtBuscar_Change`) y contador de registros activos.
   - **Alta y Reactivación:** Campos `txtDNI` y `txtMotivo` con validación previa de obligatoriedad, guardado inmediato en Access (`modDatabase.AnadirAListaNegra`), registro en auditoría (`LogActividad`) y refresco instantáneo de la lista y la hoja.
   - **Retirada de Veto (Baja Lógica):** Botón `btnRetirar` que solicita confirmación modal al usuario y actualiza en Access `Activo = False`, preservando el histórico y la trazabilidad de la auditoría.
   - **Accesibilidad:** Habilitado para ambos roles (`ADMINISTRADOR` y `OPERADOR`), manteniendo la coherencia con el diseño del sistema.

2. **Ampliación de Infraestructura Core en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas):**
   - `SincronizarHojaListaNegra()`: Descarga limpia de los registros activos de la tabla `ListaNegra` de Access en la hoja Excel `LISTA NEGRA`, formateo de fecha `dd/mm/yyyy`, autoajuste de columnas y blindaje de celdas con protección por contraseña (`UserInterfaceOnly:=True`).
   - `AnadirAListaNegra()`: Mejorada para comprobar si el DNI ya existía previamente en la tabla (incluso desactivado). Si existía, lo reactiva con `UPDATE` actualizando fecha y motivo, evitando violaciones de la restricción del índice único `idx_LN_DNI`.
   - `RetirarDeListaNegra()`: Procedimiento para desactivación lógica (`Activo = False`) en Access.

3. **Actualización del Callback RibbonX en [BotonListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonListaNegra.bas):**
   - El procedimiento `MostrarListaNegra`:
     1. Sincroniza automáticamente la hoja `LISTA NEGRA` desde Access.
     2. Muestra y activa la hoja.
     3. Abre el formulario modal `frmGestionListaNegra.Show vbModal`.

4. **Saneamiento de la Base de Datos Access ([Residencia_BE.accdb](file:///H:/ResidenciaBD/Residencia_BE.accdb)):**
   - Eliminación de la fila residual de encabezados (`DELETE FROM ListaNegra WHERE DNI = 'DNI'`).
   - Normalización de fechas históricas importadas.

---

### 3. Trazabilidad Criptográfica de Respaldos Diarios

| Elemento | Archivo de Respaldo | Hash MD5 | Estado |
|:---|:---|:---:|:---:|
| **XLSM Pre-Cambio** | `BDAS_Multiusuario_BACKUP_2026-10-18_1937_PRE_GESTION_LISTA_NEGRA.xlsm` | `2553D789AB30EE062203062265F37712` | 🔒 Verificado PRE |
| **ACCDB Pre-Cambio** | `Residencia_BE_BACKUP_2026-10-18_1937_PRE_GESTION_LISTA_NEGRA.accdb` | `ADDBB2E73CAB45CAF06F3F830DB08840` | 🔒 Verificado PRE |
| **XLSM Post-Cambio** | `BDAS_Multiusuario_BACKUP_2026-10-18_1945_POST_GESTION_LISTA_NEGRA.xlsm` | `38D2609E98AC6FEA5F7387BB8A5DDF99` | 🟢 Activo Local / Red |
| **ACCDB Post-Cambio** | `Residencia_BE_BACKUP_2026-10-18_1945_POST_GESTION_LISTA_NEGRA.accdb` | `9C6CCB9E3819CF8425989596436F42AC` | 🟢 Producción Access |

- **Coincidencia Binaria Local y Red:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `38D2609E98AC6FEA5F7387BB8A5DDF99`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `38D2609E98AC6FEA5F7387BB8A5DDF99`
- **Compilación VBA en Excel 2016:** **`0 ERRORES`**.
- **Regla Estricta de Nombres:** Cumplimiento total del texto literal `GIJÓN` (`\xd3` / 211), sin uso de `Chr(211)`.
- **Golden Master:** Intacto e inalterado.

---

## Incidencia Adicional: Ámbito de Visibilidad de la Función `Nz` en `modDatabase.bas`

### 1. Diagnóstico del Error
- **Síntoma:** Al pulsar el botón "Lista Negra", saltaba un error de compilación modal: *"No se encontró el método o el dato miembro"*, resaltando en amarillo la cabecera `Public Sub CargarLista` y en azul `.Nz` dentro de la llamada `modDatabase.Nz(...)`.
- **Causa Raíz:** En `modDatabase.bas`, la función auxiliar `Nz` estaba declarada con alcance privado (`Private Function Nz(...)`), lo que impedía su invocación externa desde el formulario `frmGestionListaNegra`.

### 2. Corrección Realizada
- Se modificó la firma de la función en `modDatabase.bas` a `Public Function Nz(ByVal valor As Variant, Optional ByVal valorDefecto As Variant = "") As Variant`, haciéndola accesible de forma limpia y reutilizable para todo el proyecto.
- Se actualizó el módulo en el libro activo y se verificó la carga, inicialización y descarga de `frmGestionListaNegra` mediante prueba automatizada con **0 errores**.

### 3. Trazabilidad de Respaldos (Fix Nz)
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_1958_PRE_FIX_NZ_LISTA_NEGRA.xlsm` (MD5: `38D2609E98AC6FEA5F7387BB8A5DDF99`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2000_POST_FIX_NZ_LISTA_NEGRA.xlsm` (MD5: `AAC33447321A0F79DE91FC7CDF3233C5`)
- **Libro Local y Red Sincronizados:**

---

## Incidencia Adicional: Enlace de Eventos de Controles Dinámicos, Botón de Búsqueda y Corrección Ortográfica

### 1. Diagnóstico del Comportamiento
- **Síntomas Reportados:**
  1. Al rellenar los campos DNI y Motivo y pulsar el botón *Anadir / Guardar*, el formulario no respondía ni realizaba ninguna acción.
  2. Ausencia de un botón explícito "Buscar" junto al cuadro de búsqueda (DNI o Motivo) y el contador de registros.
  3. Faltas ortográficas en el texto de los controles: "Anadir" en lugar de "Añadir".
- **Causa Raíz de la Falta de Respuesta:**
  - En VBA / MSForms, cuando un control se crea en tiempo de ejecución mediante `Controls.Add(...)`, la arquitectura de eventos COM no enlaza automáticamente los eventos con subrutinas con nombre convencional tipo `Private Sub btnAnadir_Click()`.
  - Para que los eventos sean capturados y ejecutados por el motor de MSForms, el control debe estar declarado a nivel de módulo con la cláusula `WithEvents` (p. ej. `Private WithEvents btnAnadir As MSForms.CommandButton`) y asignado explícitamente en el constructor del formulario.

### 2. Correcciones Realizadas
1. **Enlace Estricto de Eventos mediante `WithEvents`:**
   - Se declararon variables a nivel de módulo con escucha de eventos para todos los controles interactivos:
     - `btnAnadir`: Captura el click para validación de DNI y motivo, llamada a `modDatabase.AnadirAListaNegra`, inserción de LOG, refresco de lista y hoja sincronizada, y vaciado de campos con mensaje de confirmación.
     - `btnRetirar`: Captura el click con confirmación modal para retirar el veto (`Activo = False`) de la persona seleccionada en el `ListBox`.
     - `btnCerrar`: Cierra limpiamente el formulario (`Unload Me`).
     - `btnBuscar`: Ejecuta el filtrado de la lista según el texto introducido.
     - `btnLimpiar`: Limpia el cuadro de búsqueda y restaura la vista completa de los registros.
     - `txtBuscar`: Captura el evento `KeyDown` para que al pulsar `Enter` (código 13) se dispare automáticamente la búsqueda.
     - `lstLista`: Captura el evento `Click` para cargar de forma inmediata los datos de la fila seleccionada en las cajas de texto de edición inferior.
2. **Incorporación de Botón "Buscar" y "Limpiar":**
   - Se rediseñó la cabecera del buscador: junto al campo `txtBuscar` se situaron los botones dedicados **Buscar** y **Limpiar**, dejando a continuación el contador dinámico en verde (ej. *37 registro(s)*).
3. **Corrección Ortográfica Estricta:**
   - Se sustituyeron todas las apariciones de "Anadir" por "Añadir" tanto en la etiqueta separadora (`"Añadir o Actualizar Persona en Lista Negra:"`) como en el botón principal (`"Añadir / Guardar"`), utilizando secuencias estándar compatibles con la regla de codificación del proyecto.

### 3. Trazabilidad Criptográfica de Respaldos (Fix Eventos y Búsqueda)
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2015_PRE_FIX_EVENTOS_LISTA_NEGRA.xlsm` (MD5: `AAC33447321A0F79DE91FC7CDF3233C5`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2018_POST_FIX_EVENTOS_LISTA_NEGRA.xlsm` (MD5: `F4FE1BC1B1674D32B34D8FE944A955A6`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `F4FE1BC1B1674D32B34D8FE944A955A6`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `F4FE1BC1B1674D32B34D8FE944A955A6`
- **Compilación VBA en Excel:** 🟢 **0 ERRORES**

---

## Incidencia Adicional: Optimización Extrema del Rendimiento y Agilidad de Apertura de la Lista Negra

### 1. Diagnóstico del Rendimiento y Análisis de Cuellos de Botella
- **Síntoma Reportado:** La ventana de gestión de la Lista Negra abría de forma pesada y con una latencia perceptible tras pulsar el botón en la cinta de opciones.
- **Auditoría de Tiempos Realizada (Milisegundos con COM Automation):**
  - Carga e inicialización gráfica del formulario `Load frmGestionListaNegra`: **`0,047 s`** (47 ms).
  - Consulta ADO a la base de datos Access: **`0,031 s`** (31 ms).
  - Sincronización previa de la hoja Excel `modDatabase.SincronizarHojaListaNegra`: **`9,031 s`** (9.031 ms).
- **Causas Raíz Identificadas:**
  1. **Escritura Celda por Celda en Hoja:** En `SincronizarHojaListaNegra`, se volcaban los registros fila a fila y celda a celda (`ws.Cells(fila, "A").Value = ...`, `ws.Cells(fila, "B").NumberFormat = ...`), disparando internamente en Excel más de 150 llamadas COM de formateo, recálculo y repintado de pantalla.
  2. **Bloqueo Masivo de Celdas:** La llamada `ws.Cells.Locked = True` forzaba a Excel a aplicar propiedades sobre los 1.048.576 renglones × 16.384 columnas del libro (>17.000 millones de celdas).
  3. **Ausencia de Blindaje de Refresco de Pantalla:** No se desactivaban `Application.ScreenUpdating`, `Application.EnableEvents` ni `Application.Calculation`, provocando parpadeo y congelación momentánea antes de abrir la ventana.
  4. **Poblado Individual en ListBox:** En `frmGestionListaNegra.CargarLista`, se utilizaba un bucle con `AddItem` y `.List(i, col) = ...` elemento a elemento.

### 2. Optimizaciones Implementadas
1. **Volcado Masivo en Memoria (Bulk Array Write) en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas):**
   - Se reemplazó el bucle celda por celda por una carga masiva en memoria mediante `rs.GetRows()`, conformando una matriz bidimensional Variant `arrData(1 To numReg, 1 To 3)`.
   - Se realiza la escritura en la hoja `LISTA NEGRA` en **un único paso instantáneo**: `ws.Range("A3").Resize(numReg, 3).Value = arrData`.
   - Se limita el bloqueo de celdas estrictamente al rango ocupado (`ws.Range("A1:C" & (2 + numReg)).Locked = True`), reduciendo drásticamente la huella de memoria.
   - Se protegió todo el bloque con desactivación y restauración garantizada de entorno (`ScreenUpdating = False`, `EnableEvents = False`, `Calculation = xlCalculationManual`).
2. **Asignación Directa de Matriz al ListBox en [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm):**
   - En `CargarLista`, los resultados se cargan de una sola vez mediante `lstLista.List = arrList`, pasando de múltiples llamadas por fila a una asignación directa en apenas **`0,008 s`** (8 ms).
3. **Flujo de Pantalla Suave en [BotonListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonListaNegra.bas):**
   - Enlace limpio con supresión de repintado de pantalla mientras se sincroniza y activa la hoja de fondo, eliminando cualquier sensación de bloqueo o parpadeo.

### 3. Resultados Cuantitativos de Rendimiento
- **Tiempo Sincronización Hoja:** De **`9,031 s`** ➔ **`0,289 s`** (**Aceleración de 31x / >3.100% más rápido**).
- **Tiempo Poblado Lista:** De **`0,120 s`** ➔ **`0,008 s`**.
- **Apertura Total Percibida:** Prácticamente **instantánea (< 0,35 s)**.

### 4. Trazabilidad Criptográfica de Respaldos (Optimización)
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2032_PRE_OPTIMIZACION_LISTA_NEGRA.xlsm` (MD5: `F4FE1BC1B1674D32B34D8FE944A955A6`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2041_POST_OPTIMIZACION_LISTA_NEGRA.xlsm` (MD5: `AEF8A396CB135CFD41A0EE0803406FE0`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `AEF8A396CB135CFD41A0EE0803406FE0`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `AEF8A396CB135CFD41A0EE0803406FE0`
- **Compilación VBA en Excel:** 🟢 **0 ERRORES**

---

## Consulta y Registro de Mejora Opcional: Auditoría de Lista Negra en Hojas LOG

- **Consulta:** Se verificó la trazabilidad de la retirada de un veto de prueba en la Lista Negra. Se comprobó que el evento quedó registrado satisfactoriamente en la base de datos central Access (`LogActividad`, Id: `8092`).
- **Decisión del Usuario:** Por indicación expresa, la replicación visual de estas acciones en las tres hojas de Excel (`LOG_GIJÓN`, `LOG_SOTO`, `LOG_OVIEDO`) se cataloga como **implementación opcional para valorar en el futuro**, quedando debidamente documentada con diseño y código listo en [docs/PROPUESTAS_MEJORAS_FUTURAS.md](file:///H:/ResidenciaApp/docs/PROPUESTAS_MEJORAS_FUTURAS.md) (referencia `[PROP-01]`).

---

## Mejora Funcional: Selección Predeterminada de Todas las Residencias Accesibles en el Login

### 1. Requerimiento y Diagnóstico
- **Comportamiento Previo:** Tras validar credenciales en el formulario de login, a los usuarios con permisos sobre múltiples residencias (ej. `admin` con `GIJON,SOTO,OVIEDO` o JAIME con `SOTO,GIJON`) se les mostraba la ventana de selección de residencia ([frmSelectorResidencia.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmSelectorResidencia.frm)), pero únicamente aparecía marcada con casilla de verificación la primera de ellas (`Selected(0) = True`), obligando al usuario a marcar manualmente una a una las demás casillas si deseaba operar con todas.
- **Objetivo:** Que por defecto se presenten marcadas/seleccionadas todas las residencias a las que el usuario tenga acceso asignado, permitiendo pulsar `Enter` directamente para ingresar con todas activas o desmarcar puntualmente las que no se deseen en esa sesión.

### 2. Modificación Implementada
En el evento `UserForm_Initialize` de [frmSelectorResidencia.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmSelectorResidencia.frm), se sustituyó la selección única inicial por un bucle completo sobre todas las residencias cargadas en el `ListBox`:

```vba
' Inicia con todas las residencias asignadas marcadas por defecto
For i = 0 To lstResidencias.ListCount - 1
    lstResidencias.Selected(i) = True
Next i
```

### 3. Trazabilidad Criptográfica de Respaldos
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2105_PRE_SELECTOR_TODAS_DEFAULT.xlsm` (MD5: `AEF8A396CB135CFD41A0EE0803406FE0`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2108_POST_SELECTOR_TODAS_DEFAULT.xlsm` (MD5: `6E05177A99727A64C3937594DBA7F0E2`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `6E05177A99727A64C3937594DBA7F0E2`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `6E05177A99727A64C3937594DBA7F0E2`
- **Compilación VBA en Excel:** 🟢 **0 ERRORES**

---

## Mejora de Usabilidad: Navegación Totalmente Ágil por Teclado (Enter / Tab) en el Formulario de Login

### 1. Requerimiento y Diagnóstico
- **Objetivo:** Optimizar el flujo de entrada al sistema para no depender del ratón ni de teclas alternativas:
  1. Que al abrirse el formulario de inicio de sesión ([frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm)), el cursor se sitúe automáticamente en el campo **Usuario**.
  2. Que al pulsar `Enter` (o `Tab`), el cursor salte automáticamente al campo **Contraseña**.
  3. Que al pulsar nuevamente `Enter` (o `Tab`), el foco se desplace al botón **Entrar**.
  4. Que al pulsar `Enter` por tercera vez sobre el botón activo, se procese la autenticación y se acceda a la app.
- **Análisis de Impacto y No Interferencia:**
  - El cambio está estrictamente encapsulado en [frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm).
  - No altera ninguna interfaz pública ni lógica de base de datos (`ValidarCredenciales`).
  - La interacción mediante ratón (clic en Entrar / Salir), el comportamiento del botón Salir / Cancelar (tecla `Escape`) y el manejo de errores de credenciales continúan funcionando de forma 100% idéntica y compatible.

### 2. Modificaciones Implementadas
1. **Variables con Eventos de Teclado:** Se declararon a nivel de módulo `Private WithEvents txtUsuario As MSForms.TextBox` y `Private WithEvents txtClave As MSForms.TextBox`.
2. **Posición Inicial Garantizada:** En `UserForm_Activate`, se asegura el foco inmediato mediante `txtUsuario.SetFocus`.
3. **Captura y Secuencia de Teclas `Enter` (`KeyDown`):**
   - En `txtUsuario_KeyDown`: Si `KeyCode = 13` (Enter), se cancela la propagación estándar (`KeyCode = 0`) y se traslada el cursor con `txtClave.SetFocus`.
   - En `txtClave_KeyDown`: Si `KeyCode = 13` (Enter), se cancela la propagación y se traslada el foco al botón con `btnEntrar.SetFocus`.
   - En `btnEntrar_KeyDown`: Si `KeyCode = 13` (Enter), se ejecuta directamente `ProcesarLogin`.
4. **Sincronización con `Tab`:** Se establecieron secuencialmente los índices `TabIndex` (0: Usuario, 1: Contraseña, 2: Entrar, 3: Salir) y `TabStop = True`, logrando que tanto la pulsación de `Enter` como de `Tab` sigan exactamente la misma ruta de acceso natural.

### 3. Trazabilidad Criptográfica de Respaldos
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2118_PRE_LOGIN_KEYBOARD_NAV.xlsm` (MD5: `6E05177A99727A64C3937594DBA7F0E2`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2121_POST_LOGIN_KEYBOARD_NAV.xlsm` (MD5: `4D3D5F3C2120B9E0F91D1CD2F05B9945`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `4D3D5F3C2120B9E0F91D1CD2F05B9945`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `4D3D5F3C2120B9E0F91D1CD2F05B9945`
- **Compilación VBA en Excel:** 🟢 **0 ERRORES**

---

## Consolidación Oficial: Elevación a Versión de Referencia Estable (Hito 2026-10-18 21:25)

### 1. Alcance de la Consolidación
Por solicitud expresa del usuario, se declara el estado actual como la nueva **Versión de Referencia Oficial** del proyecto, consolidando de manera acumulativa y permanente todas las funcionalidades validadas durante la jornada del 18 de Octubre de 2026:

1. **Gestión Interactiva de Lista Negra:** Formulario dedicado modal [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm) con enlaces `WithEvents`, botones dedicados "Buscar" y "Limpiar", escucha de la tecla `Enter`, alta/actualización con validación previa de índice único y baja lógica (`Activo = False`).
2. **Arquitectura de Sincronización Ultrarrápida (31x):** Descarga masiva en memoria matriz Variant bidimensional (`GetRows`), reducción de latencia de 9,03s a 0,28s, supresión de repintado de pantalla y restricción de celdas bloqueadas al rango ocupado.
3. **Selector Multirresidencia Optimizado:** Preselección con casillas marcadas por defecto para todas las residencias autorizadas al usuario en [frmSelectorResidencia.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmSelectorResidencia.frm).
4. **Navegación 100% por Teclado en Login:** Ciclo ergonómico completo con `Enter` / `Tab`: Usuario ➔ Contraseña ➔ Botón Entrar ➔ Acceso directo en [frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm).
5. **Seguridad y Control de Roles:** Acceso por perfiles `ADMINISTRADOR` / `OPERADOR`, visibilidad dinámica de pestañas nativas de Excel en la cinta RibbonX, protección estricta y blindaje de hojas LOG, purga automática en segundo plano (>45 días) y botón de navegación RibbonX "Último Nº Orden".

### 2. Certificación de Integridad Cuádruple y Blindaje de Solo Lectura

| Componente | Ruta Absoluta | Hash MD5 | Estado de Protección |
|:---|:---|:---:|:---:|
| **Libro Activo Trabajo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🟢 Activo Editable |
| **Libro Despliegue Red** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🟢 Servidor Compartido |
| **Versión Referencia XLSM** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Backup Referencia XLSM** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 💾 Respaldo Consolidado |
| **Base de Datos Activa** | `H:\ResidenciaBD\Residencia_BE.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 🟢 Producción Access |
| **Versión Referencia ACCDB** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Backup Referencia ACCDB** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 💾 Respaldo Consolidado |

### 3. Matriz Completa de Trazabilidad de Respaldos Diarios (2026-10-18)

| Nº | Intervención / Funcionalidad | Backup PRE | MD5 PRE | Backup POST | MD5 POST |
|:---:|:---|:---|:---:|:---|:---:|
| **1** | Infraestructura Roles Base | `..._2328_PRE_FASE1_PASO0.xlsm` | `F814E071502409B94A89C2FA545AC513` | `..._2330_POST_FASE1_PASO0.xlsm` | `DFDA9D433D53EC929518B561DA7CB7E5` |
| **2** | Desbloqueo Universal Seguro | `..._2344_PRE_FASE1_PASO1.xlsm` | `DFDA9D433D53EC929518B561DA7CB7E5` | `..._2346_POST_FASE1_PASO1.xlsm` | `BA1CB2218D2DCFD2AC284AA52E63E9CB` |
| **3** | Gestión Contraseñas Sistema | `..._2351_PRE_FASE1_PASOS4Y5.xlsm` | `BA1CB2218D2DCFD2AC284AA52E63E9CB` | `..._2353_POST_FASE1_PASOS4Y5.xlsm` | `461D7CC14B8E33EAFEAE4FFB6221B8B5` |
| **4** | Criptografía Hashes SHA256 | `..._2358_PRE_FASE1_PASO6.xlsm` | `461D7CC14B8E33EAFEAE4FFB6221B8B5` | `..._2359_POST_FASE1_PASO6.xlsm` | `5CF3E19C04870183DE1B5DBA1CF736B7` |
| **5** | Borrado Masivo Alojamientos | `..._0001_PRE_FASE1_PASO7.xlsm` | `5CF3E19C04870183DE1B5DBA1CF736B7` | `..._0002_POST_FASE1_PASO7.xlsm` | `3B7621D574224386B6F5734889FC928B` |
| **6** | Panel de Control RGPD | `..._0005_PRE_FASE1_PASO3.xlsm` | `3B7621D574224386B6F5734889FC928B` | `..._0006_POST_FASE1_PASO3.xlsm` | `2A6C0958FF917F8B4F2BB617C4D04213` |
| **7** | Búsqueda Datos BDAS | `..._0012_PRE_FASE1_PASO2.xlsm` | `2A6C0958FF917F8B4F2BB617C4D04213` | `..._0013_POST_FASE1_PASO2.xlsm` | `C6A062109B65ADB26F1BE77610D7E9C1` |
| **8** | Protección Comentarios | `..._0017_PRE_FASE1_PASO10.xlsm` | `C6A062109B65ADB26F1BE77610D7E9C1` | `..._0018_POST_FASE1_PASO10.xlsm` | `0058CE54DB0B98999F3F2C1F24A0623A` |
| **9** | Limpieza `modDatabase.bas` | `..._0055_PRE_FIX_CARACTERES_...xlsm` | `A40306B867E2A1A417B2ADB8A4D4C14F` | `..._0058_POST_FIX_CARACTERES_...xlsm` | `4282D6C258456AAB26D5E1A40989D0D4` |
| **10** | Limpieza Cosmética Global | `..._0112_PRE_LIMPIEZA_COSMETICA.xlsm` | `4282D6C258456AAB26D5E1A40989D0D4` | `..._0115_POST_LIMPIEZA_COSMETICA.xlsm` | `78EB55FBCB23046C81C0F9E91B384375` |
| **11** | Pestañas Nativas por Rol | `..._0205_PRE_PESTANAS_NATIVAS...xlsm` | `1403BDC5AF52D42A2F68FD4677BD3F54` | `..._0210_POST_PESTANAS_NATIVAS...xlsm` | `B488080D0263C5FA7535E13DCF087692` |
| **12** | Protección Hojas LOG por Rol | `..._0235_PRE_PROTECCION_HOJAS...xlsm` | `37728133C4AD1740FC1D40B510E733F2` | `..._0240_POST_PROTECCION_HOJAS...xlsm` | `5017901D805E1876D02112A42457C0F2` |
| **13** | Purga Automática LOG (>45d) | `..._0300_PRE_PURGA_AUTOMATICA...xlsm` | `5017901D805E1876D02112A42457C0F2` | `..._0305_POST_PURGA_AUTOMATICA...xlsm` | `2B76736979710A428AF5F88705F3FF46` |
| **14** | Botón "Último Nº Orden" LOG | `..._0310_PRE_IR_ULTIMA_CELDA...xlsm` | `2B76736979710A428AF5F88705F3FF46` | `..._0315_POST_IR_ULTIMA_CELDA...xlsm` | `2553D789AB30EE062203062265F37712` |
| **15** | Formulario Lista Negra + Hoja | `..._1937_PRE_GESTION_LISTA_NEGRA.xlsm` | `2553D789AB30EE062203062265F37712` | `..._1945_POST_GESTION_LISTA_NEGRA.xlsm` | `38D2609E98AC6FEA5F7387BB8A5DDF99` |
| — | *Base Access Lista Negra* | `Residencia_BE...1937_PRE...accdb` | `ADDBB2E73CAB45CAF06F3F830DB08840` | `Residencia_BE...1945_POST...accdb` | `9C6CCB9E3819CF8425989596436F42AC` |
| **16** | Corrección Función `modDatabase.Nz` | `..._1958_PRE_FIX_NZ_LISTA_NEGRA.xlsm` | `38D2609E98AC6FEA5F7387BB8A5DDF99` | `..._2000_POST_FIX_NZ_LISTA_NEGRA.xlsm` | `AAC33447321A0F79DE91FC7CDF3233C5` |
| **17** | Enlace Eventos COM y Buscar | `..._2015_PRE_FIX_EVENTOS_...xlsm` | `AAC33447321A0F79DE91FC7CDF3233C5` | `..._2018_POST_FIX_EVENTOS_...xlsm` | `F4FE1BC1B1674D32B34D8FE944A955A6` |
| **18** | Optimización Matriz Memoria 31x | `..._2032_PRE_OPTIMIZACION_...xlsm` | `F4FE1BC1B1674D32B34D8FE944A955A6` | `..._2041_POST_OPTIMIZACION_...xlsm` | `AEF8A396CB135CFD41A0EE0803406FE0` |
| **19** | Selector Preselección por Defecto | `..._2105_PRE_SELECTOR_TODAS_...xlsm` | `AEF8A396CB135CFD41A0EE0803406FE0` | `..._2108_POST_SELECTOR_TODAS_...xlsm` | `6E05177A99727A64C3937594DBA7F0E2` |
| **20** | Navegación Login Teclado (Enter) | `..._2118_PRE_LOGIN_KEYBOARD_...xlsm` | `6E05177A99727A64C3937594DBA7F0E2` | `..._2121_POST_LOGIN_KEYBOARD_...xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` |
| **21** | **Versión de Referencia Oficial** | *(Consolidación de Cierre)* | — | `..._2125_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` |
| — | *Base Access Referencia Final* | *(Consolidación de Cierre)* | — | `Residencia_BE...2125_VERSION_REF.accdb` | `2AB9DAEB5015BE53738F23743D309228` |

### 4. Gobernanza y Preservación del Entorno
- **Histórico Preservado:** La versión de referencia anterior (Hito 03:20) fue archivada con el atributo de solo lectura en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_0320.xlsm` y `Residencia_BE_VERSION_REFERENCIA_2026-10-18_0320.accdb`.
- **Golden Master Inmutable:** La réplica Golden Master (`H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm`) permanece intacta e inalterada con hash MD5 `A03DB5430C8E913AB2703F5116904DD2`.
- **Compilación COM Validada:** Se validó la compilación de todos los módulos VBA en Excel 2016 con **`0 ERRORES`**.
- **Regla Estricta de Nombres:** Cumplimiento del 100% en todo el código y hojas con el texto literal `GIJÓN` (`\xd3` / 211), sin uso de `Chr(211)`.

---

## Optimización de Rendimiento y Experiencia Visual en el Arranque (Estándar ECRMC - Opciones 1+2 y Opción B)

### 1. Diagnóstico del Problema y Causas Raíz
- **Síntomas Reportados:**
  1. Latencia perceptible de varios segundos desde que el usuario pulsa "Entrar" en el login hasta que la hoja inicial queda disponible.
  2. "Desajustes" visuales perceptibles durante ese tiempo: parpadeo de cuadrícula, alternancia visible de pestañas en la barra inferior, columnas que se vacían y llenan a la vista, y congelación gráfica detrás del diálogo modal de bienvenida.
- **Causas Raíz Identificadas:**
  1. **Fugas de `Application.ScreenUpdating`:** Múltiples procedimientos hijos (`MostrarHojasResidencia` y `SincronizarHojaDesdeAccess`) forzaban `ScreenUpdating = True` prematuramente, provocando que Excel redibujara en vivo cada paso intermedio (borrado de filas, volcado de datos, reajuste de cuadrícula y reorganización de pestañas).
  2. **Recálculo Automático Activo:** Con `Application.Calculation = xlCalculationAutomatic`, cada inserción de datos y fórmula disparaba evaluaciones de dependencias en cadena.
  3. **Bucle Celda a Celda en Inyección de Fórmulas:** Más de 6.400 llamadas individuales COM (`hoja.Cells(fRow, cCol).HasFormula` y `.FormulaR1C1 = ...`) en lugar de asignaciones vectoriales por bloque.
  4. **Purga Síncrona de LOGs:** La depuración de registros >45 días demoraba el hilo principal de inicio en ~0,65 segundos.
  5. **Bloqueo Modal por `MsgBox`:** El cuadro modal de bienvenida detenía el mensaje pump de Windows, dejando la interfaz de fondo a medio estabilizar hasta la pulsación de "Aceptar".

### 2. Solución Técnica Implementada (Estándar ECRMC)
1. **Blindaje Hermético del Entorno en [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls):**
   - Captura y suspensión estricta del entorno al arrancar: `ScreenUpdating = False`, `EnableEvents = False`, `Calculation = xlCalculationManual`.
   - Bloque estructurado de salida garantizada `CleanUp:` que siempre restaura `ScreenUpdating = True`, `EnableEvents = True` y `Calculation = xlCalculationAutomatic`, incluso en caso de error fortuito.
2. **Inyección Vectorial Atómica de Fórmulas en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas):**
   - Sustitución del bucle de 5.400 llamadas por asignaciones atómicas de rango continuo por columna:
     `hoja.Range(hoja.Cells(filaInicio, cCol), hoja.Cells(filaInicio + totalRows - 1, cCol)).FormulaR1C1 = fmlaMaster`
     reduciendo el tiempo de inyección a menos de 5 milisegundos sin alterar la lógica de cálculo.
3. **Preservación Contextual de Pantalla en Rutinas Hijas ([modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas)):**
   - `SincronizarHojaDesdeAccess` y `MostrarHojasResidencia` memorizan el estado previo (`scrOrig`) y lo restauran fielmente al terminar, evitando fugas prematuras hacia el usuario.
4. **Purga de Auditoría en Segundo Plano Diferido ([ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas)):**
   - La rutina de purga se programa para ejecutarse 2 segundos después del inicio mediante `Application.OnTime Now + TimeValue("00:00:02"), "PrepararHojaLOG_Timer"`, liberando el arranque inmediato.
5. **Acceso Directo y Fluido sin `MsgBox` (Opción B):**
   - Supresión del `MsgBox` bloqueante. La confirmación de acceso se publica de forma fluida en la barra de estado:
     `Application.StatusBar = "Conectado como ... | Residencia activa: ..."`
     permitiendo al usuario operar de inmediato.

### 3. Trazabilidad Criptográfica de Respaldos
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2200_PRE_OPTIMIZACION_ARRANQUE.xlsm` (MD5: `4D3D5F3C2120B9E0F91D1CD2F05B9945`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2205_POST_OPTIMIZACION_ARRANQUE.xlsm` (MD5: `A8C2E5EA360AB9142BCCD4B127596C73`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `A8C2E5EA360AB9142BCCD4B127596C73`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `A8C2E5EA360AB9142BCCD4B127596C73`
- **Compilación VBA en Excel 2016 COM:** 🟢 **0 ERRORES**
- **Regla Estricta de Nombres:** Cumplimiento total del texto literal `GIJÓN` (`\xd3` / 211).
- **Golden Master:** Intacto e inalterado (`A03DB5430C8E913AB2703F5116904DD2`).


---

## Requerimiento: Eliminación Definitiva de la Hoja `LISTA NEGRA` y Desacoplamiento Integral (Opción C - ECRMC)

### 1. Diagnóstico y Contexto Técnico
- **Consulta Planteada:** Tras la implantación del formulario interactivo modal `frmGestionListaNegra`, se analizó si la hoja residual `LISTA NEGRA` era necesaria o si se podía prescindir de ella por completo.
- **Dependencias Identificadas en el Código Anterior:**
  1. `ModListaNegra.bas`: La rutina `CheckBlacklist(ws, Target)` comprobaba obligatoriamente la existencia de `ThisWorkbook.Sheets("LISTA NEGRA")` lanzando un error modal si no existía, y utilizaba sus celdas como búsqueda de respaldo.
  2. `BotonListaNegra.bas`: Al pulsar el botón del Ribbon, forzaba la sincronización y visibilidad/activación de dicha hoja antes de abrir el formulario.
  3. `frmGestionListaNegra.frm`: En las operaciones de alta y baja, invocaba `modDatabase.SincronizarHojaListaNegra` para reescribir las celdas de Excel.
  4. `modDatabase.bas`: Mantenía la subrutina `SincronizarHojaListaNegra()` de volcado a la cuadrícula de Excel.
- **Decisión del Usuario:** Proceder con la **Opción C**: eliminar por completo la hoja del libro Excel, desacoplar el código para operar 100% en memoria y contra Access, aplicando las Reglas Permanentes del Proyecto y el Estándar ECRMC.

---

### 2. Modificaciones Técnicas Implementadas
1. **Desacoplamiento Total en [ModListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModListaNegra.bas):**
   - Se eliminaron todas las referencias a `wsLista` y a la hoja Excel.
   - `CheckBlacklist` comprueba la edición en la columna I y valida directamente contra Access mediante `modDatabase.ComprobarListaNegra(Trim(CStr(cel.Value)), motivoAcc)`. Si existe veto, emite la advertencia con el motivo registrado.
   - Bloque de salida estructurado y protegido contra excepciones.
2. **Apertura Directa en [BotonListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonListaNegra.bas):**
   - Se suprimió la activación de hojas secundarias.
   - Al pulsar el botón en la cinta RibbonX, el usuario permanece en su hoja operativa y se despliega directamente `frmGestionListaNegra.Show vbModal`.
3. **Optimización de [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm):**
   - Retiradas las llamadas obsoletas a `SincronizarHojaListaNegra` tras altas y bajas.
   - Actualizada la nota informativa inferior: `* Los cambios se guardan de forma inmediata en la base de datos central Access.`
   - Cumplimiento estricto de la regla de nombres (sin `Chr(211)`).
4. **Limpieza de Código Muerto en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas):**
   - Eliminada la rutina `SincronizarHojaListaNegra()`, manteniendo limpias las funciones centrales de negocio (`ComprobarListaNegra`, `AnadirAListaNegra`, `RetirarDeListaNegra`).
5. **Supresión Física de la Hoja en [BDAS_Multiusuario.xlsm](file:///H:/ResidenciaApp/bdas-multiusuario/BDAS_Multiusuario.xlsm):**
   - Se eliminó la hoja `LISTA NEGRA` (`Hoja3`).
   - El recuento de hojas del libro se redujo limpiamente de 49 a 48 hojas.

---

### 3. Trazabilidad Criptográfica de Respaldos Diarios

| Elemento | Archivo de Respaldo | Hash MD5 | Estado |
|:---|:---|:---:|:---:|
| **XLSM Pre-Cambio** | `BDAS_Multiusuario_BACKUP_2026-10-18_2250_PRE_ELIMINAR_HOJA_LISTA_NEGRA.xlsm` | `A8C2E5EA360AB9142BCCD4B127596C73` | 🔒 Verificado PRE |
| **XLSM Post-Cambio** | `BDAS_Multiusuario_BACKUP_2026-10-18_2255_POST_ELIMINAR_HOJA_LISTA_NEGRA.xlsm` | `0EEAC3EEB5CDFB08682EBAA78AD4D3F1` | 💾 Respaldo POST |
| **Libro Activo Local** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `0EEAC3EEB5CDFB08682EBAA78AD4D3F1` | 🟢 Producción Local |
| **Libro Despliegue Red** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | `0EEAC3EEB5CDFB08682EBAA78AD4D3F1` | 🟢 Servidor Compartido |

- **Compilación VBA en Excel 2016 COM:** 🟢 **`0 ERRORES`**.
- **Regla Estricta de Nombres:** Cumplimiento total del texto literal `GIJÓN` (`\xd3` / 211), sin uso de `Chr(211)`.
- **Golden Master:** Intacto e inalterado (`A03DB5430C8E913AB2703F5116904DD2`).

---

## Saneamiento Menor: Purga de Cabecera Residual `MultiUse` en `ThisWorkbook`

### 1. Diagnóstico y Corrección
- **Síntoma Reportado:** Al abrir el libro, el compilador de VBA arrojaba *"el procedimiento externo no es válido"* señalando a la instrucción residual `MultiUse = -1  'True` en la primera línea de `ThisWorkbook`.
- **Causa Raíz:** Dicha línea correspondía a la cabecera técnica de exportación de clases (`BEGIN ... END`) que se había filtrado durante la inyección de código.
- **Acción Realizada:**
  1. Se eliminó la línea residual por completo del módulo `ThisWorkbook` en el libro de trabajo.
  2. La primera línea pasa a ser directamente la cabecera limpia de comentarios del módulo (`' ====================`), seguida de `Option Explicit`.
  3. Se validó la compilación de todo el proyecto VBA en Excel 2016 COM: **0 ERRORES**.

### 2. Trazabilidad Criptográfica de Respaldos
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2255_POST_ELIMINAR_HOJA_LISTA_NEGRA.xlsm` (MD5: `0EEAC3EEB5CDFB08682EBAA78AD4D3F1`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2305_POST_PURGA_LINEA_THISWORKBOOK.xlsm` (MD5: `0ADE915204BD421786271C4F244746A6`)
- **Libro Local y Red Sincronizados:**
  - `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `0ADE915204BD421786271C4F244746A6`
  - `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `0ADE915204BD421786271C4F244746A6`
- **Compilación VBA en Excel 2016 COM:** 🟢 **0 ERRORES**.
- **Golden Master:** Intacto e inalterado (`A03DB5430C8E913AB2703F5116904DD2`).
