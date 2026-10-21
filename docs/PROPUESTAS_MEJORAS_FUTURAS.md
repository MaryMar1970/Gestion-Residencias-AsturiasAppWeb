# Catálogo de Propuestas y Mejoras Opcionales para Valoración Futura

Este documento recopila propuestas técnicas, extensiones funcionales y mejoras identificadas durante el ciclo de vida del proyecto que quedan reservadas como **opcionales** para su análisis y decisión en sesiones futuras, respetando las prioridades y directrices del usuario.

---

## Índice de Propuestas

| ID | Fecha Registro | Módulo / Área | Descripción Resumida | Estado |
|:---:|:---:|:---|:---|:---:|
| **PROP-01** | 18/10/2026 | Auditoría / Lista Negra | Replicación de Altas y Bajas de Lista Negra en Hojas LOG de Excel | 📋 Resuelta (Sync desde Access) |
| **PROP-02** | 19/10/2026 | Interfaz / Arranque | Mitigación y Amortiguación de Parpadeos Visuales (Flicker) al Iniciar la App y Login | ⏳ Pendiente de Revisión |
| **PROP-03** | 21/10/2026 | Resiliencia / Conectividad | Sistema de Continuidad Operativa Fuera de Línea (Modo Contingencia Nivel 2 — "Buffer & Sync") | 📋 Planificada para Próxima Fase |

---

## Fichas de Detalle Técnico

### [PROP-01] Replicación de Altas y Bajas de Lista Negra en las Tres Hojas LOG de Excel

- **Fecha de Propuesta:** 18 de octubre de 2026.
- **Origen:** Consulta sobre la no aparición de la retirada de un veto en las hojas `LOG_GIJÓN`, `LOG_SOTO` y `LOG_OVIEDO`.
- **Estado:** 📋 **Registrada como implementación opcional para valorar en el futuro.**

#### 1. Diagnóstico del Estado Actual
- Actualmente, la gestión de la Lista Negra ([frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm)) registra todas las operaciones de inclusión y retirada de personas en la base de datos central Access (`Residencia_BE.accdb`, tabla `LogActividad`) de forma permanente e inmutable a través de `modDatabase.InsertarLog`.
- Las tres hojas de Excel (`LOG_GIJÓN`, `LOG_SOTO` y `LOG_OVIEDO`), accesibles desde el Ribbon de la cinta de opciones (**AUDITORÍA**), son hojas de log locales asociadas a cada residencia individual. Dichas hojas registran cambios de celdas en las hojas de órdenes, inicio de sesión (LOGIN) y cambios en el estado de habitaciones.
- Al ser la Lista Negra una entidad global y transversal a todas las residencias, el formulario no vuelca actualmente apuntes en las tres hojas locales de Excel.

#### 2. Propuesta Técnica Lista para Implementar
Crear en [ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas) un procedimiento específico:

```vba
'=================================================================================
' PROCEDIMIENTO: Registrar acciones de Lista Negra en las hojas LOG de Excel
' Parametros:
'   - usuario: Nombre del usuario autenticado (ej: "admin")
'   - accion: "ALTA_LISTA_NEGRA" o "BAJA_LISTA_NEGRA"
'   - dni: DNI de la persona afectada
'   - motivo: Motivo del veto o detalle descriptivo
'=================================================================================
Public Sub RegistrarListaNegraLOG(ByVal usuario As String, _
                                  ByVal accion As String, _
                                  ByVal dni As String, _
                                  Optional ByVal motivo As String = "")
    Dim logPassword As String
    Dim j As Long
    Dim wsLOG As Worksheet
    Dim filaLog As Long
    
    InicializarNombresHojasLOG
    logPassword = ModuloConfigSegura.ObtenerPasswordLOG()
    
    ' La Lista Negra es transversal: se registra de forma unificada en las tres hojas
    For j = LBound(nombresHojasLOG) To UBound(nombresHojasLOG)
        Set wsLOG = Nothing
        On Error Resume Next
        Set wsLOG = ThisWorkbook.Worksheets(nombresHojasLOG(j))
        On Error GoTo 0
        
        If Not wsLOG Is Nothing Then
            On Error Resume Next
            filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
            wsLOG.Cells(filaLog, 1).Value = usuario
            wsLOG.Cells(filaLog, 2).Value = Now
            wsLOG.Cells(filaLog, 2).NumberFormat = "dd/mm/yyyy hh:mm"
            wsLOG.Cells(filaLog, 3).Value = accion
            wsLOG.Cells(filaLog, 4).Value = "DNI: " & dni
            wsLOG.Cells(filaLog, 5).Value = IIf(accion = "BAJA_LISTA_NEGRA", "VETO RETIRADO", "VETO ACTIVO")
            wsLOG.Cells(filaLog, 6).Value = motivo
            
            ' Fallback con desproteccion y reproteccion si la hoja estuviese protegida
            If Err.Number <> 0 Then
                Err.Clear
                wsLOG.Unprotect password:=logPassword
                filaLog = wsLOG.Cells(wsLOG.Rows.Count, "A").End(xlUp).Row + 1
                wsLOG.Cells(filaLog, 1).Value = usuario
                wsLOG.Cells(filaLog, 2).Value = Now
                wsLOG.Cells(filaLog, 2).NumberFormat = "dd/mm/yyyy hh:mm"
                wsLOG.Cells(filaLog, 3).Value = accion
                wsLOG.Cells(filaLog, 4).Value = "DNI: " & dni
                wsLOG.Cells(filaLog, 5).Value = IIf(accion = "BAJA_LISTA_NEGRA", "VETO RETIRADO", "VETO ACTIVO")
                wsLOG.Cells(filaLog, 6).Value = motivo
                wsLOG.Protect password:=logPassword, UserInterfaceOnly:=True, AllowFiltering:=True, AllowSorting:=True
            End If
            On Error GoTo 0
            Set wsLOG = Nothing
        End If
    Next j
End Sub
```

Y su invocación directa en [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm):
- En `btnAnadir_Click`: `ModuloLOG.RegistrarListaNegraLOG modDatabase.ObtenerUsuarioActual(), "ALTA_LISTA_NEGRA", dni, motivo`
- En `btnRetirar_Click`: `ModuloLOG.RegistrarListaNegraLOG modDatabase.ObtenerUsuarioActual(), "BAJA_LISTA_NEGRA", dni, motivo`

#### 3. Criterios de Valoración para el Futuro
- **Ventajas:**
  - Los usuarios que auditan las acciones exclusivamente abriendo los botones del Ribbon (**`LOG Gijón`**, **`LOG Soto`**, **`LOG Oviedo`**) podrán consultar de inmediato los movimientos de la Lista Negra sin necesidad de acceder directamente a la base de datos Access.
- **Alternativas a valorar:**
  - Opción A: Replicar en las tres hojas existentes (como está diseñado arriba).
  - Opción B: Crear un botón o visor específico en el Ribbon para "LOG Global / Auditoría Central" que lea directamente de la tabla `LogActividad` de Access.

---

### [PROP-02] Mitigación y Amortiguación de Parpadeos Visuales (Flicker) y Re-renderizado de Carátula al Iniciar la App y Login

- **Fecha de Registro:** 19 de octubre de 2026.
- **Origen:** Observación y petición expresa del usuario tras comprobar que, al abrir la aplicación e introducir usuario/contraseña, la imagen de fondo se reinicia a tramos.
- **Estado:** ⏳ **Anotada y Priorizada para FUTURA REVISIÓN según instrucción expresa del usuario.**
- **Módulos Implicados:**
  - Hoja `INICIO` (Forma gráfica incrustada `ImgCaratulaCasaCuartel`)
  - [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) (`Workbook_Open`, `Workbook_BeforeClose`)
  - [frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm) (`UserForm_Initialize`, `UserForm_Activate`, `btnEntrar_Click`)
  - [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (`MostrarHojasResidencia`, `SincronizarTodasMisResidencias`)

#### 1. Descripción de la Incidencia / Comportamiento Observado
- Al abrir el libro de trabajo `.xlsm`, la aplicación carga la hoja `INICIO` con la fotografía de la Casa Cuartel de fondo y presenta el formulario modal `frmLogin`.
- Tanto durante la apertura inicial de la aplicación como tras introducir usuario/contraseña y cerrarse el formulario, la imagen de la carátula no se dibuja en un único bloque limpio, sino que **se reinicia a tramos** (parpadeo por franjas o redibujado secuencial de bandas horizontales/tiles por parte del motor gráfico de Excel GDI+/DirectX).
- Este defecto visual afecta a la experiencia estética del arranque, percibiendo el usuario que la imagen se recompone a trompicones antes de dar paso a la hoja operativa.

#### 2. Diagnóstico Técnico de la Causa Raíz
1. **Renderizado de Formas Gráficas en el Canvas de Hojas Excel**: A diferencia de una ventana nativa de Windows o un formulario UserForm, los objetos `Shape` (tipo 13 `msoPicture`) incrustados en celdas de Excel son dibujados por el subsistema gráfico de la cuadrícula. Al superponerse un formulario modal (`frmLogin`) o al invalidarse el área cliente (`WM_PAINT`), Excel descompone el repintado de bitmaps de alta resolución (1570x845 px escalados a 1177x634 pt) en bandas de exploración (*scanline tiles*).
2. **Ciclo de Enfoque y Desaparición del Cuadro Modal**: Al pulsar "Entrar", `frmLogin` se oculta (`Me.Hide` o `Unload`), dejando un "hueco sucio" en la memoria de vídeo exactamente donde estaba el formulario. Excel redibuja primero esa porción central de la imagen y posteriormente el resto del canvas, produciendo el efecto visual de "reinicio a tramos".
3. **Interacción con ScreenUpdating y Recálculo de Fondo**: Si `Application.ScreenUpdating` conmuta o si se ejecutan eventos de activación de hojas o cálculo mientras la imagen está en el búfer de pantalla, el motor de aceleración por hardware de Office redibuja selectivamente las regiones dañadas a ritmos desiguales.

#### 3. Vías Técnicas de Solución para la Futura Revisión
1. **Vía A (Recomendada): Carátula Integrada en UserForm a Pantalla Completa (Splash Screen Nativo)**:
   - En lugar de insertar la fotografía como una forma sobre las celdas de una hoja Excel (`INICIO`), cargar la imagen directamente dentro de un control `Image` en un formulario UserForm maximizado/sin bordes o Splash Screen integrado.
   - Los formularios VBA de Microsoft Forms utilizan un búfer de ventana dedicado de Windows que no se descompone en bandas de celdas, garantizando un renderizado 100% monolítico e instantáneo.
   - El formulario de Login puede integrarse directamente embebido en la misma pantalla o cerrarse conjuntamente sin dejar parches de repintado.
2. **Vía B: Optimización de Resolución y Compresión del Bitmap**:
   - Ajustar la resolución exacta de la imagen al tamaño en puntos de la visualización (`96 DPI` a `1178x634 px` en lugar de los `1570x845 px` actuales), reduciendo el peso en memoria gráfica y permitiendo que un único ciclo de copia de memoria traslade el mapa de bits sin necesidad de particionado en franjas.
3. **Vía C: Congelación a Nivel de Ventana del Sistema Operativo (Win32 API `LockWindowUpdate` o `WM_SETREDRAW`)**:
   - Interceptar el manejador de la ventana de Excel (`Application.Hwnd`):
     `LockWindowUpdate Application.Hwnd`
   - Bloquear el envío de mensajes de repintado a la GPU mientras se procesa la validación de credenciales, y desbloquear (`LockWindowUpdate 0`) únicamente cuando la hoja de residencia final esté activada, evitando que el usuario vea la fase intermedia de recomposición de la imagen.
4. **Vía D: Ajuste de Aceleración Gráfica por Hardware**:
   - Analizar el impacto de la aceleración gráfica por hardware de Office (`DisableHardwareAcceleration` en el registro o directiva de aplicación) en el equipo cliente.

---

---

### [PROP-03] Sistema de Continuidad Operativa Fuera de Línea (Modo Contingencia Nivel 2 — "Buffer & Sync")

- **Nombre del Proyecto de Mejora:** **SICO-N2** (*Sistema de Continuidad Operativa Nivel 2*) / **Operativa de Contingencia "Buffer & Sync"**.
- **Fecha de Registro y Planificación:** 21–22 de octubre de 2026.
- **Origen:** Análisis y decisión expresa del usuario para dotar a la aplicación de resiliencia total ante pérdidas prolongadas de comunicación con la base de datos central (`H:\ResidenciaBD\Residencia_BE.accdb`).
- **Estado:** 📋 **Planificada y Priorizada para Próxima Fase de Trabajo.**
- **Módulos Implicados:**
  - Hoja de Sistema Oculta `SYS_COLA_SYNC` (Persistencia no volátil en disco de operaciones locales).
  - [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (Capa de persistencia ADO, comprobación rápida de red y motor de volcado FIFO).
  - `ModuloSeguridadLocal.bas` (Caché local cifrada de credenciales para login desconectado).
  - [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) (Control de ciclo de vida, arranque en contingencia, salvaguarda en cierre).
  - Hojas operativas: `Hoja2` (`RESIDENCIA GIJÓN`), `Hoja3` (`RESIDENCIA SOTO DEL BARCO`), `Hoja4` (`RESIDENCIA OVIEDO`) en sus eventos `Worksheet_Change`.
  - Barra de Estado (`Application.StatusBar`) y Ribbon/Cabecera (Semáforo de Red e indicador de cambios en búfer).

---

#### 1. Objetivos Estratégicos del Proyecto
1. **Cero Bloqueos al Iniciar:** Permitir que el operador encienda su ordenador y comience a trabajar de inmediato, incluso si la unidad `H:\` o la red están completamente caídas desde primera hora de la mañana.
2. **Capacidad Operativa Real (Recepción Activa):** Despachar todas las operaciones ordinarias de recepción sin necesidad de apuntes en papel (cobros en metálico/banco, actualización de teléfonos, matrículas, notas de estancia, check-in).
3. **Persistencia Permanente Multidiaria:** Los datos acumulados no se pierden al apagar el equipo por la tarde; se conservan físicamente en el archivo `.xlsm` durante días o fines de semana completos hasta que regrese la conexión.
4. **Sincronización Invisible y Reconciliación FIFO:** Al regresar la red, el sistema vuelca automáticamente todas las operaciones en orden cronológico estricto a Access en menos de 2 segundos, preservando la fecha y hora original de cada cobro para la auditoría.
5. **Blindaje de Líneas Rojas Contables:** Bloqueo preventivo de numeraciones correlativas de facturas y asignación de habitaciones simultáneas no consolidadas para evitar colisiones fiscales o duplicidades entre residencias.

---

#### 2. Arquitectura Técnica de la Solución

```
+-----------------------------------------------------------------------------------+
|                            FRONT-END (Excel .xlsm)                                |
|                                                                                   |
|  [Operador edita celda / registra cobro]                                          |
|                       |                                                           |
|                       v                                                           |
|             ¿Hay conexión con H:\?                                                |
|            /                      \                                              |
|         (SÍ)                      (NO)                                            |
|          |                         |                                              |
|          v                         v                                              |
|   Persistencia ADODB      [SYS_COLA_SYNC] (Hoja oculta en disco)                  |
|   directa en Access       +---------------------------------------------+         |
|   (Residencia_BE.accdb)   | ID | FechaHora | Usuario | Orden | Campo | Valor|         |
|                           +---------------------------------------------+         |
|                                    |                                              |
|                                    | (Persiste apagados y reinicios)              |
|                                    v                                              |
|                         [Al detectar retorno de red]                              |
|                                    |                                              |
|                                    v                                              |
|                       Sincronización Diferida FIFO                                |
|                       -> Vuelca cambios a Access con fecha real                   |
|                       -> Registra apunte en LogActividad                         |
|                       -> Vacía la cola y emite notificación                      |
+-----------------------------------------------------------------------------------+
```

---

#### 3. Especificación de Componentes a Desarrollar

##### A. Hoja de Sistema Oculta: `SYS_COLA_SYNC`
- Propiedad: `xlSheetVeryHidden` (invisible en el selector de hojas de Excel).
- Protegida con la clave maestra de hojas (`ModuloConfigSegura.ObtenerPasswordHojas()`).
- Columnas normalizadas:
  1. `ID_Transaccion` (Autonumérico local GUID/Long).
  2. `FechaHoraLocal` (Fecha y hora exacta en que el operador realizó la acción).
  3. `Usuario` (Identificador del operador que registró el cambio).
  4. `Residencia` (`GIJON`, `SOTO`, `OVIEDO`).
  5. `NumOrden` (Nº de orden de la reserva afectada).
  6. `TipoOperacion` (`UPDATE_PAGO`, `UPDATE_CONTACTO`, `UPDATE_ESTADO`, `PRERESERVA`).
  7. `CampoModificado` (Nombre del campo en Access: `EstadoPago`, `Telefono`, `Direccion`, etc.).
  8. `ValorAnterior` (Para trazabilidad y rollback).
  9. `ValorNuevo` (Valor que debe consolidarse en Access).
  10. `EstadoSync` (`PENDIENTE`, `SINCRONIZANDO`, `CONCILIADO`, `ERROR`).
  11. `Reintentos` (Contador de intentos).

##### B. Módulo de Caché de Credenciales Locales: `ModuloSeguridadLocal.bas`
- En cada login exitoso conectado a Access, guarda en un Nombre Definido oculto de Excel (`app_local_cred_cache`) el par `[Usuario | Hash_SHA256 | Residencias | Timestamp]`, cifrado con sal local vinculada a la máquina.
- Al arrancar sin red:
  * Si la conexión falla, `Workbook_Open` no se cierra; invoca el login de contingencia.
  * Valida las credenciales contra la caché local.
  * Si coinciden, abre las hojas de la residencia correspondiente en **Modo Contingencia Nivel 2**.

##### C. Semáforo de Red y Monitor No Bloqueante
- Indicador visual permanente en la Barra de Estado de Excel:
  * 🟢 **`[CONECTADO - H:\ResidenciaBD]`**: Operativa ordinaria en vivo.
  * 🟡 **`[MODO CONTINGENCIA N2 - X cambios pendientes]`**: Desconectado; datos asegurados en búfer local.
  * 🔄 **`[SINCRONIZANDO CON SERVIDOR...]`**: Volcando cola en segundo plano.
- Detección de fallo ultra-rápida: si una petición da error de red, activa el modo desconectado durante 60 segundos antes del siguiente sondeo, eliminando retardos o relojes de arena en la cuadrícula.

##### D. Enrutador de Guardado en Hojas Operativas (`Hoja2`, `Hoja3`, `Hoja4`)
- En `Worksheet_Change`, si el semáforo está en amarillo/rojo:
  * Desvía el cambio inmediatamente a `SYS_COLA_SYNC` mediante el procedimiento `EncolarOperacionOffline`.
  * La celda de Excel se actualiza en pantalla al instante (0 milisegundos de espera).
  * Opcional: marca visual sutil (borde tenue o sombreado suave) indicando que ese dato específico está pendiente de subir a Access.

##### E. Motor de Reconciliación Automática (FIFO)
- Procedimiento `SincronizarColaOfflineHaciaAccess`:
  * Se dispara al detectar el paso de 🔴 a 🟢 (o mediante botón en cinta de opciones: *"Sincronizar Cola Ahora"*).
  * Ordena los registros pendientes por `ID_Transaccion` ascendente (cronológico estricto).
  * Ejecuta los `UPDATE` transaccionales sobre la tabla `Ordenes` en Access.
  * Inserta en `LogActividad` de Access el registro de auditoría con la fecha y hora real en que ocurrió la acción offline.
  * Elimina las filas de `SYS_COLA_SYNC` solo tras confirmación positiva de Access.
  * Muestra resumen al usuario: *"🟢 Sincronización completada: X operaciones consolidadas en el servidor central"*.

##### F. Blindaje de Líneas Rojas Contables y de Asignación
- **Facturación:** El botón de asignación de número de factura muestra un aviso informativo:
  > *"La emisión de números oficiales de factura requiere conexión con el servidor central para garantizar la correlatividad estricta. Esta función se activará en cuanto se restablezca la red."*
- **Nuevas Reservas:** Si el operador introduce una nueva orden sin red, se registra con estado `"PRE-RESERVA EN ESPERA"`. Al volver la red, el sistema comprueba la no duplicidad y disponibilidad de la habitación antes de consolidarla como orden firme.

---

#### 4. Plan de Implementación en 5 Fases

| Fase | Tareas y Alcance | Entregables / Módulos |
| :---: | :---| :---|
| **Fase 1** | **Estructura y Persistencia del Búfer Local:**<br>- Creación de la hoja oculta protegida `SYS_COLA_SYNC`.<br>- Procedimientos `EncolarOperacionOffline` y lectura de cola.<br>- Prueba de persistencia al guardar y reabrir el archivo `.xlsm`. | `SYS_COLA_SYNC`<br>`modDatabase.bas` |
| **Fase 2** | **Autenticación Desconectada en Caché:**<br>- Cifrado y almacenamiento local de credenciales.<br>- Adaptación de `frmLogin.frm` y `Workbook_Open` para tolerar arranque sin red.<br>- Validación segura de contingencia y apertura de hojas asignadas. | `ModuloSeguridadLocal.bas`<br>`frmLogin.frm`<br>`ThisWorkbook.cls` |
| **Fase 3** | **Semáforo de Conectividad y Enrutador de Celdas:**<br>- Implementación del indicador visual en `StatusBar`.<br>- Intercepción transparente en `Worksheet_Change` de Gijón, Soto y Oviedo.<br>- Eliminación de retardos y bloqueos de red. | `Hoja2`, `Hoja3`, `Hoja4`<br>`ThisWorkbook.cls` |
| **Fase 4** | **Motor de Sincronización Diferida (FIFO):**<br>- Procedimiento de volcado masivo cronológico a Access.<br>- Registro en `LogActividad` con marca temporal histórica.<br>- Notificaciones al usuario y vaciado seguro de la cola. | `modDatabase.bas`<br>`ModuloLOG.bas` |
| **Fase 5** | **Blindaje de Líneas Rojas y Pruebas de Estrés:**<br>- Protección de facturas y pre-reservas.<br>- Simulación de corte de 24/48 horas con apagado de equipo.<br>- Verificación forense de cuadre en Access tras reconexión. | Protocolo de Pruebas<br>Documentación Final |

---

#### 5. Criterios de Calidad y Gobernanza
- Cumplimiento estricto del estándar **ECRMC** (código modular, tipado explícito, control de errores hermético).
- Protocolo **Cero Mojibake** (Windows-1252 ANSI nativo).
- **Inmutabilidad del Golden Master** y generación de copias de seguridad incrementales antes de cada fase.
