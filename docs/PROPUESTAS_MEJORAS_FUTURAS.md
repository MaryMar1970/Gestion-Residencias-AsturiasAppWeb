# Catálogo de Propuestas y Mejoras Opcionales para Valoración Futura

Este documento recopila propuestas técnicas, extensiones funcionales y mejoras identificadas durante el ciclo de vida del proyecto que quedan reservadas como **opcionales** para su análisis y decisión en sesiones futuras, respetando las prioridades y directrices del usuario.

---

## Índice de Propuestas

| ID | Fecha Registro | Módulo / Área | Descripción Resumida | Estado |
|:---:|:---:|:---|:---|:---:|
| **PROP-01** | 18/10/2026 | Auditoría / Lista Negra | Replicación de Altas y Bajas de Lista Negra en Hojas LOG de Excel | 📋 Resuelta (Sync desde Access) |
| **PROP-02** | 19/10/2026 | Interfaz / Arranque | Mitigación y Amortiguación de Parpadeos Visuales (Flicker) al Iniciar la App y Login | ⏳ Pendiente de Revisión |
| **PROP-03** | 21/10/2026 | Conectividad / BD | Monitor de Conectividad en Tiempo Real ("Semáforo de Red") y Detección Temprana de Pérdida de Conexión con Access | ⏳ Pendiente de Valoración |

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

### [PROP-03] Monitor de Conectividad en Tiempo Real ("Semáforo de Red") y Detección Temprana de Pérdida de Conexión con el Backend Access

- **Fecha de Registro:** 21 de octubre de 2026.
- **Origen:** Consulta del usuario sobre la resiliencia y el comportamiento del Front-End (Excel) ante caídas, cortes o degradación de la red hacia el Back-End (`H:\ResidenciaBD\Residencia_BE.accdb`).
- **Estado:** ⏳ **Pendiente de Valoración.**
- **Módulos Implicados:**
  - [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (`ComprobarConexion`, `SincronizarFilaHaciaAccess`, `ActualizarOrden`, `ExecuteNonQuery`)
  - [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) / Barra de Estado (`Application.StatusBar`) / Encabezado de Hoja
  - Hojas operativas (`Hoja2`, `Hoja3`, `Hoja4` - `Worksheet_Change`)

#### 1. Diagnóstico del Estado Actual y Justificación
- La arquitectura actual de BDAS Multiusuario está diseñada bajo un modelo cliente-servidor transaccional directo en tiempo real.
- Al arrancar, si no hay conexión con `H:\ResidenciaBD\Residencia_BE.accdb`, la aplicación bloquea el acceso de forma segura y se cierra para evitar trabajar a ciegas sin autenticación ni sincronización.
- Sin embargo, si la conexión se interrumpe o degrada **con la aplicación ya abierta**:
  1. Excel mantiene los datos en la memoria RAM local y permite al usuario seguir introduciendo o modificando valores en la cuadrícula.
  2. Cada edición dispara llamadas ADODB que intentan conectar con Access; si la red falla o está degradada, se consumen hasta 5 segundos de espera (`ConnectionTimeout`) por celda, provocando congelaciones intermitentes.
  3. Los cambios editados quedan en la pantalla local pero **no se graban en la base de datos central**.
  4. Si el operador no advierte el fallo y sigue trabajando, al reabrir la aplicación en la siguiente sesión los datos descargados desde Access sobrescribirán lo tecleado fuera de línea, provocando pérdida involuntaria de trabajo.

#### 2. Vías Técnicas de Solución Propuestas
1. **Vía A: Indicador Visual Permanente de Estado ("Semáforo de Red")**:
   - Integrar un indicador de latencia y estado en la Barra de Estado de Excel (`Application.StatusBar`) o como un pequeño elemento visual en la cabecera fija de las hojas de residencia:
     * 🟢 **`[CONECTADO - H:\ResidenciaBD]`**: Comunicación fluida con el backend Access.
     * 🟡 **`[CONEXIÓN LENTA]`**: Tiempo de respuesta superior a 2 segundos en la última consulta.
     * 🔴 **`[SIN RED / DESCONECTADO]`**: Base de datos inaccesible; advertencia inmediata al usuario.
2. **Vía B: Alerta Temprana Emergente y Protección Preventiva**:
   - Modificar las rutinas de sincronización en caliente (`SincronizarFilaHaciaAccess` y `ActualizarOrden`) para que, ante el primer fallo de escritura en Access:
     * Intercepte el error antes de propagar más llamadas a la red.
     * Muestre un aviso claro e instructivo:
       > *"ALERTA DE CONEXIÓN: Se ha interrumpido la comunicación con la base de datos central en H:\ResidenciaBD. Los últimos cambios no se han podido guardar. Por favor, detén la introducción de datos y comprueba el acceso a la red antes de continuar."*
     * Evita que el operador continúe introduciendo datos a ciegas creyendo que se están guardando.
3. **Vía C: Búfer Temporal de Reintentos en Memoria (Cola Local Ligera)**:
   - Para microcortes de pocos segundos, retener las filas modificadas en una colección interna (`FilasPendientesReintento`) y reintentar su inserción/actualización de forma transparente en cuanto la conexión vuelva a estar disponible, evitando la pérdida de datos ante pequeñas fluctuaciones de red.

#### 3. Criterios de Valoración
- **Ventajas:**
  - Transparencia total para los operadores sobre la salud y disponibilidad de su conexión de red.
  - Prevención garantizada de pérdidas de datos por desconexión no advertida.
  - Experiencia de usuario (UX) mucho más sólida, profesional y confiable.
- **Riesgos / Impacto:** Mínimo. No altera la estructura de tablas de Access ni los cálculos contables o de facturación de las hojas.
