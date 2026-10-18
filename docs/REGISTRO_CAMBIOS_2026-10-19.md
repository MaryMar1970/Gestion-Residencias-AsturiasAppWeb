# Registro Diario de Control de Cambios — 19/10/2026

> **Entorno de Trabajo Activo:** `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
> **Servidor Compartido en Red:** `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
> **Base de Datos Backend:** `H:\ResidenciaBD\Residencia_BE.accdb`  
> **Versión de Referencia Oficial:** `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (Hito 2026-10-19 01:50)  
> **Estado de Compilación VBA:** 🟢 **0 ERRORES** (Validado en Microsoft Excel 2016 COM)

---

## 1. Resumen Ejecutivo de la Sesión

En esta sesión se abordaron e implementaron con éxito siete grandes hitos funcionales y correctivos, cumpliendo estrictamente con el **Estándar de Calidad, Robustez y Mantenibilidad del Código (ECRMC)**, las **Reglas Permanentes del Proyecto** y el protocolo de copias de seguridad incrementales:

1. **Eliminación Definitiva de la Hoja `LISTA NEGRA` de Excel (Opción C)**: Migración hacia un modelo 100% en memoria con origen único en la base de datos Access.
2. **Corrección de Sintaxis de Cabecera en `ThisWorkbook.cls`**: Supresión de la instrucción inválida `MultiUse = -1 'True` generada por herramientas externas de exportación.
3. **Resolución de Error Crítico en Persistencia de Reservas y LOGs**: Creación de rutinas de guardado atómico (UPSERT ADODB) en tiempo real para evitar la pérdida de reservas al navegar entre residencias y calendarios, junto con sincronización masiva de auditoría LOG desde Access.
4. **Blindaje de Visibilidad de las Hojas de Gijón**: Corrección de codificación binaria en la comparación de cadenas (`GIJÓN` / `GIJON`), resolviendo la ocultación no deseada tras salir del calendario.
5. **Carátula Institucional en Hoja `INICIO` a Pantalla Completa (Opción 1)**: Configuración de la hoja de presentación inicial con cuadrículas y encabezados ocultos, bloqueo de scroll y centrado modal de `frmLogin`.
6. **Sustitución de Imagen por Fotografía Histórica de la Casa Cuartel**: Inserción y embebido nativo permanente de [Foto Casa Cuartel antigua (3).png](file:///H:/ResidenciaApp/bdas-multiusuario/Foto%20Casa%20Cuartel%20antigua%20(3).png).
7. **Asignación Automática por Defecto de `"SIN TURNO"` en Residencia Soto**: Configuración para que al asignarse una evaluación en la columna G, se asigne automáticamente `"SIN TURNO"` en la columna H manteniendo total libertad para su modificación posterior por parte del usuario.

---

## 2. Detalle de Incidencias y Requerimientos Abordados

### 2.1. Eliminación Definitiva de la Hoja `LISTA NEGRA` (Opción C)
- **Problema previo:** Existencia de una hoja visible `LISTA NEGRA` en el libro de trabajo que requería mantenimiento redundante y podía desalinearse de Access.
- **Acción implementada:**
  - Supresión física y definitiva de la hoja `LISTA NEGRA` del archivo `.xlsm`.
  - Desacople integral del código VBA: se actualizaron [BotonListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonListaNegra.bas), [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm), [ModuloConfigSegura.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloConfigSegura.bas) y [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas).
  - La verificación de lista negra en las hojas operativas de residencia se ejecuta en memoria mediante un `Scripting.Dictionary` indexado por DNI, cargado en milisegundos directamente desde la tabla `ListaNegra` de Access.

### 2.2. Corrección de Cabecera en `ThisWorkbook.cls`
- **Error detectado:** Al abrir el libro se mostraba el error de compilación: `El procedimiento externo no es válido: MultiUse = -1 'True`.
- **Causa:** Exportación e importación de la clase que dejó expuesta la directiva interna de VBA fuera del bloque de atributos.
- **Acción implementada:** Se depuró la cabecera del módulo de clase [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls), dejándolo estrictamente estructurado según la especificación de clases VBA de Microsoft Excel.

### 2.3. Corrección de Error Crítico: Persistencia de Reservas y LOGs
- **Incidencia crítica:** Al introducir los datos de una reserva en `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO` o `RESIDENCIA OVIEDO`:
  1. Los datos de las reservas no se trasladaban a los calendarios.
  2. Al regresar desde el calendario a la hoja de residencia, los datos introducidos habían desaparecido (no se guardaban en Access).
  3. No se registraban en las hojas de LOG de auditoría.
- **Causa raíz:** Las celdas modificadas en las hojas de residencia no tenían disparadores automáticos de guardado relacional hacia Access antes de cambiar de hoja. Al navegar hacia el calendario o cerrar la hoja, `Workbook_BeforeClose` (`ThisWorkbook.Saved = True`) o el refresco de vista limpiaba las filas no consolidadas.
- **Solución técnica implementada:**
  1. En [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas), creación de `AsegurarOrdenExiste(numOrden, residencia)` para garantizar que el registro maestro exista en la tabla `Ordenes` con su residencia e identificador primario.
  2. Creación de `SincronizarFilaHaciaAccess(ws, fila)`: función de sincronización relacional (UPSERT) mediante ADODB que lee todos los campos de la fila modificada (Nº Factura, Finalidad, Empleo, Situación, Evaluación, Turno, DNI, Rango, Nombre, Fechas Entrada/Salida, Días de Uso, Habitación/Apartamento, Importes, Observaciones) y ejecuta un `UPDATE` parametrizado e inmediato sobre la base de datos Access.
  3. Enlace en los eventos `Worksheet_Change` de las tres residencias, garantizando que al editar cualquier dato de una reserva, la fila se guarde atómicamente.
  4. Enlace en los formularios de selección de habitación ([SeleccionHabitacionGijon.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/SeleccionHabitacionGijon.frm), `SeleccionHabitacionSOTO.frm`, `SeleccionHabitacionOVIEDO.frm`) para persistir la asignación antes de cerrar el diálogo.
  5. Sincronización de Auditoría LOG: Creación de `ModuloLOG.SincronizarHojaLOGDesdeAccess` en [ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas) para descargar automáticamente los últimos 500 registros de la tabla `LogActividad` de Access formateados en las hojas de log de Excel.

### 2.4. Blindaje de Visibilidad de las Hojas de Gijón
- **Incidencia detectada:** Al acceder a las tres residencias y editar en Gijón, al navegar a `Calendario GIJÓN` y salir de él, desaparecían u ocultaban las hojas relacionadas con Gijón (`RESIDENCIA GIJÓN`, `RESUMEN GIJÓN` y `Calendario GIJÓN`), mientras que en Soto y Oviedo no ocurría.
- **Causa:** En [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) (`Workbook_SheetDeactivate`) y [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (`MostrarHojasResidencia`), la cadena `"GIJÓN"` contenía bytes UTF-8 (`\xc3\x93`) que no coincidían con el nombre ANSI interno de la hoja de Excel (`\xd3`), provocando que la rutina clasificara la hoja como secundaria y la ocultara (`Sh.Visible = xlSheetHidden`).
- **Solución implementada:** Normalización y evaluación dual en `Workbook_SheetDeactivate` y `MostrarHojasResidencia`, reconociendo tanto `GIJÓN` como `GIJON`. Se probó la navegación y se confirmó que las tres hojas de Gijón permanecen 100% visibles.

### 2.5. Pantalla de Inicio con Carátula de Presentación a Pantalla Completa (Opción 1)
- **Requerimiento:** Al inicio de la carga de la aplicación, mostrar una carátula limpia de presentación de fondo a pantalla completa hasta que el usuario introduzca sus credenciales.
- **Solución implementada:**
  - Hoja `INICIO` rediseñada: cuadrículas (`DisplayGridlines = False`) y encabezados de fila/columna (`DisplayHeadings = False`) desactivados, fondo blanco puro (`RGB(255, 255, 255)`), área de desplazamiento restringida (`ScrollArea = "A1:BA45"`) y hoja protegida (`DrawingObjects:=True`).
  - Inserción y embebido permanente (`SaveWithDocument:=True`) de la imagen.
  - El libro se guarda con `INICIO` como hoja visible y activa, y las hojas operativas ocultas. Al hacer doble clic sobre el archivo `.xlsm`, Excel renderiza instantáneamente la carátula antes de iniciar la ejecución de macros.
  - En `Workbook_Open`, `frmLogin` emerge centrado sobre la carátula.
  - Al completar el login y selector de residencias, `modDatabase.MostrarHojasResidencia` activa la hoja de la residencia y oculta automáticamente la hoja `INICIO` (`wsInicio.Visible = xlSheetHidden`).
  - En `Workbook_BeforeClose`, `INICIO` vuelve a prepararse como activa para la siguiente apertura.

### 2.6. Sustitución de Imagen por Fotografía Histórica de la Casa Cuartel
- **Requerimiento:** Reemplazar la carátula inicial por [Foto Casa Cuartel antigua (3).png](file:///H:/ResidenciaApp/bdas-multiusuario/Foto%20Casa%20Cuartel%20antigua%20(3).png).
- **Implementación:**
  - Se eliminó la forma gráfica anterior y se incrustó la fotografía original de alta resolución (1570 x 845 px).
  - Escalado proporcional a 1177.5 x 633.75 pt con bloqueo de relación de aspecto (`LockAspectRatio = msoTrue`).
  - El formulario modal `frmLogin` queda perfectamente enmarcado en el centro de la escena entre los dos guardias civiles con capa.

### 2.7. Asignación Automática por Defecto de `"SIN TURNO"` en Residencia Soto
- **Requerimiento:** En la hoja `RESIDENCIA SOTO`, al mostrarse un valor en la columna G (`EVALUACIÓN`), mostrar por defecto en la columna H (`TURNO`) el valor `"SIN TURNO"`, permitiendo que el usuario lo cambie cuando lo desee.
- **Implementación en Hoja20:**
  - En `Worksheet_Change`:
    - En `Case 4 To 6` (cálculo de evaluación tras seleccionar Finalidad, Empleo y Situación): si la columna H está vacía, se asigna `"SIN TURNO"`.
    - En `Case 7` (evaluación directa): si la columna H está vacía, se asigna `"SIN TURNO"`.
    - Si se borra la evaluación, se limpia automáticamente la columna H si contenía `"SIN TURNO"`.
  - La celda mantiene su menú desplegable de validación de datos (`SIN TURNO;S. Santa 1..3;Navidad 1..3;Verano 1..11`). El usuario puede seleccionar libremente cualquier turno.
  - Al seleccionar un turno, `Case 8` carga las fechas correspondientes; si se regresa a `"SIN TURNO"`, se limpian las fechas fijas para entrada libre.
  - Se probó mediante simulación automatizada en la fila 50 con validación 100% exitosa.

---

## 3. Matriz de Copias de Seguridad Incrementales Diarias

De acuerdo con las reglas permanentes del proyecto, cada modificación ha quedado respaldada antes y después de su aplicación en el repositorio seguro `H:\ResidenciaBD\BackupsDiarios\`:

| Timestamp | Archivo de Copia de Seguridad | Tamaño (bytes) | Hash MD5 | Hito Funcional / Incidencia Asociada |
| :---: | :---| :---:| :---:| :---|
| **18/10 22:50** | `BDAS_Multiusuario_BACKUP_2026-10-18_2250_PRE_ELIMINAR_HOJA_LISTA_NEGRA.xlsm` | 5.917.035 | `a8c2e5ea360ab9142bccd4b127596c73` | Respaldo previo a eliminación de hoja Lista Negra |
| **18/10 22:55** | `BDAS_Multiusuario_BACKUP_2026-10-18_2255_POST_ELIMINAR_HOJA_LISTA_NEGRA.xlsm` | 5.872.905 | `0eeac3eeb5cdfb08682ebaa78ad4d3f1` | Hoja Lista Negra eliminada y desacoplada |
| **18/10 23:05** | `BDAS_Multiusuario_BACKUP_2026-10-18_2305_POST_PURGA_LINEA_THISWORKBOOK.xlsm` | 6.304.607 | `0ade915204bd421786271c4f244746a6` | Limpieza cabecera `MultiUse = -1` en ThisWorkbook |
| **18/10 23:44** | `BDAS_Multiusuario_BACKUP_2026-10-18_2344_PRE_FIX_PERSISTENCIA_RESERVAS.xlsm` | 6.304.607 | `0ade915204bd421786271c4f244746a6` | Respaldo previo a corrección de persistencia |
| **18/10 23:44** | `Residencia_BE_BACKUP_2026-10-18_2344_PRE_FIX_PERSISTENCIA_RESERVAS.accdb` | 6.524.928 | `f2b056feb19781ea34277f6fb6bd000a` | Respaldo de BD Access previo a persistencia |
| **19/10 00:15** | `BDAS_Multiusuario_BACKUP_2026-10-19_0015_POST_FIX_PERSISTENCIA_RESERVAS.xlsm` | 5.586.969 | `78d953747b222726ba82d7e5537aa725` | Persistencia ADODB UPSERT + Sync LOG implementados |
| **19/10 00:15** | `Residencia_BE_BACKUP_2026-10-19_0015_POST_FIX_PERSISTENCIA_RESERVAS.accdb` | 6.623.232 | `2fdec434e69ca05494ec238050a40d4e` | BD Access sincronizada con nuevas reservas |
| **19/10 00:45** | `BDAS_Multiusuario_BACKUP_2026-10-19_0045_OK_PERSISTENCIA_RESERVAS_Y_VISIBILIDAD.xlsm` | 5.785.541 | `31b81038fa6b14071be236b4002eeb05` | Blindaje visibilidad Gijón + comprobación reservas |
| **19/10 00:45** | `Residencia_BE_BACKUP_2026-10-19_0045_OK_PERSISTENCIA_RESERVAS_Y_VISIBILIDAD.accdb` | 6.627.328 | `3d24065d103c84058523aee24ed25a41` | BD Access verificada con pruebas Gijón/Soto/Oviedo |
| **19/10 00:58** | `BDAS_Multiusuario_BACKUP_2026-10-19_0058_PRE_CARATULA_INICIO.xlsm` | 5.785.541 | `31b81038fa6b14071be236b4002eeb05` | Respaldo previo a configuración carátula inicial |
| **19/10 01:05** | `BDAS_Multiusuario_BACKUP_2026-10-19_0105_POST_CARATULA_INICIO_OPCION1.xlsm` | 6.148.675 | `dcf5a3414f58082bd5aab46f91998077` | Carátula Opción 1 en hoja INICIO configurada |
| **19/10 01:10** | `BDAS_Multiusuario_BACKUP_2026-10-19_0110_PRE_CAMBIO_FOTO_CARATULA.xlsm` | 6.148.675 | `dcf5a3414f58082bd5aab46f91998077` | Respaldo previo a cambio de fotografía |
| **19/10 01:12** | `BDAS_Multiusuario_BACKUP_2026-10-19_0112_POST_CAMBIO_FOTO_CARATULA.xlsm` | 6.425.290 | `c20e6997c74fc99db2780d948b56730a` | Fotografía Casa Cuartel antigua embebida |
| **19/10 01:40** | `BDAS_Multiusuario_BACKUP_2026-10-19_0140_PRE_SOTO_TURNO_DEFECTO.xlsm` | 6.425.290 | `c20e6997c74fc99db2780d948b56730a` | Respaldo previo a lógica de turno en Soto |
| **19/10 01:42** | `BDAS_Multiusuario_BACKUP_2026-10-19_0142_POST_SOTO_TURNO_DEFECTO.xlsm` | 6.426.870 | `54d94f60f542de5210795c3c30a5a419` | `"SIN TURNO"` por defecto en Soto Col H completado |

---

## 4. Estado de la Versión de Referencia Oficial

Conforme a las instrucciones del usuario, la versión actual de la aplicación se ha establecido y consolidado como la **nueva versión de referencia oficial**:

- **Ubicación:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`
- **Archivos y Atributos:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (MD5: `54d94f60f542de5210795c3c30a5a419`) — **Modo Solo Lectura activado**.
  - `Residencia_BE_VERSION_REFERENCIA.accdb` (MD5: `ede9c7ecb540604f835facc39f8f08b8`) — **Modo Solo Lectura activado**.
- **Histórico Preservado:** La versión de referencia anterior (Hito 2026-10-18 21:25) ha sido preservada intacta en la subcarpeta `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`.
- **Script de Restauración:** Actualizado [Restaurar-Version-Referencia.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Version-Referencia.ps1) para restaurar este hito exacto si fuera requerido.

---

## 5. Tareas Pendientes para Próxima Sesión

Se ha registrado en [PROPUESTAS_MEJORAS_FUTURAS.md](file:///H:/ResidenciaApp/docs/PROPUESTAS_MEJORAS_FUTURAS.md) como tarea prioritaria:

- **PROP-02: Revisión y amortiguación de parpadeos visuales (`flicker`) al arrancar la app e introducir usuario/contraseña**:
  - Analizar la secuencia temporal entre `Application.ScreenUpdating = True`, el pintado de la carátula estática en `INICIO`, la renderización de los controles de `frmLogin` y la transición final hacia `MostrarHojasResidencia`.
  - Evaluar la posibilidad de prefijar los controles de `frmLogin` de forma nativa en el diseñador del UserForm en lugar de construirlos dinámicamente en tiempo de ejecución (`Controls.Add`), evitando repintados intermedios.
  - Congelar la ventana mediante API Win32 (`LockWindowUpdate`) durante la carga del formulario para eliminar cualquier parpadeo de pantalla.
