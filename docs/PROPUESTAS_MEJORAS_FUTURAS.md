# Catálogo de Propuestas y Mejoras Opcionales para Valoración Futura

Este documento recopila propuestas técnicas, extensiones funcionales y mejoras identificadas durante el ciclo de vida del proyecto que quedan reservadas como **opcionales** para su análisis y decisión en sesiones futuras, respetando las prioridades y directrices del usuario.

---

## Índice de Propuestas

| ID | Fecha Registro | Módulo / Área | Descripción Resumida | Estado |
|:---:|:---:|:---|:---|:---:|
| **PROP-01** | 18/10/2026 | Auditoría / Lista Negra | Replicación de Altas y Bajas de Lista Negra en Hojas LOG de Excel | 📋 Resuelta (Sync desde Access) |
| **PROP-02** | 19/10/2026 | Interfaz / Arranque | Mitigación y Amortiguación de Parpadeos Visuales (Flicker) al Iniciar la App y Login | ⏳ Pendiente de Revisión |

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

### [PROP-02] Mitigación y Amortiguación de Parpadeos Visuales (Flicker) al Iniciar la App y Login

- **Fecha de Registro:** 19 de octubre de 2026.
- **Origen:** Petición expresa del usuario tras desplegar la carátula de inicio a pantalla completa.
- **Estado:** ⏳ **Anotada como PENDIENTE para revisión y optimización en próxima sesión.**
- **Módulos Implicados:**
  - [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) (`Workbook_Open`)
  - [frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm) (`UserForm_Initialize`, `UserForm_Activate`, `btnEntrar_Click`)
  - [frmSelectorResidencia.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmSelectorResidencia.frm)
  - [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (`MostrarHojasResidencia`, `SincronizarTodasMisResidencias`)

#### 1. Descripción de la Incidencia / Comportamiento Observado
- Al abrir el libro de trabajo `.xlsm`, la aplicación carga la hoja `INICIO` con la carátula institucional de la Casa Cuartel y presenta el formulario modal `frmLogin`.
- Durante la transición de apertura de Excel, la inicialización de controles dinámicos de `frmLogin`, el repintado de ventana y la posterior llamada a `modDatabase.MostrarHojasResidencia` (donde se alternan visibilidades de hojas y recálculo), se perciben ligeros parpadeos visuales en pantalla (*screen flickering* o refrescos intermedios de interfaz).

#### 2. Vías Técnicas Propuestas para Análisis y Corrección
1. **Encapsulación estricta de `Application.ScreenUpdating`**:
   - Asegurar que `Application.ScreenUpdating = False` cubra exactamente las transiciones entre hojas sin interferir con el renderizado estático de `INICIO`.
   - Congelar el repintado de la ventana de Excel mediante API Win32 (`LockWindowUpdate(Application.Hwnd)`) durante la carga del UserForm si fuera necesario.
2. **Optimización de carga en `frmLogin`**:
   - Sustituir la creación dinámica en tiempo de ejecución de controles (`Controls.Add`) por controles estáticos prefijados en el diseñador de `frmLogin`, evitando ciclos innecesarios de redibujado de la ventana modal.
3. **Sincronización en segundo plano**:
   - Diferir el recálculo masivo (`Application.Calculate`) y la purga de auditoría al final absoluto de la cadena de eventos tras mostrar la hoja activa de residencia.

