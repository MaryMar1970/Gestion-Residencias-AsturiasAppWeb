# Índice Maestro — Versión de Referencia Oficial (Hito 2026-10-18 21:25)

> **Estado:** CONSOLIDADO Y CERTIFICADO  
> **Fecha de Consolidación:** 18 de Octubre de 2026, 21:25 CEST  
> **Archivo de Trabajo Activo:** `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
> **Réplica en Red Servidor:** `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
> **Versión de Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (`IsReadOnly = True`)  
> **Base de Datos Backend Activa:** `H:\ResidenciaBD\Residencia_BE.accdb`  
> **BD Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` (`IsReadOnly = True`)  
> **Huella Digital Oficial (MD5 XLSM):** `4D3D5F3C2120B9E0F91D1CD2F05B9945`  
> **Huella Digital Oficial (MD5 ACCDB):** `2AB9DAEB5015BE53738F23743D309228`  
> **Estado de Compilación VBA (Excel 2016 COM):** `0 ERRORES`  

---

## 1. Certificado Cuádruple de Integridad Criptográfica

Todos los entornos de desarrollo, producción, réplica maestra y último respaldo incremental diario presentan una coincidencia binaria absoluta al 100%:

| Ámbito | Ubicación Física del Archivo | Hash MD5 | Estado |
|:---|:---|:---|:---|
| **Trabajo Local** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🟢 Activo (Producción Local) |
| **Despliegue Red** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🟢 Servidor Compartido |
| **Versión Referencia** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Último Respaldo** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` | 💾 Backup Incremental Consolidado |
| **Base Datos Backend** | `H:\ResidenciaBD\Residencia_BE.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 🟢 Producción Access |
| **BD Referencia** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Respaldo BD Ref** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.accdb` | `2AB9DAEB5015BE53738F23743D309228` | 💾 Backup Incremental Consolidado |

---

## 2. Inventario Detallado de Funcionalidades y Vinculación a Copias Incrementales

A continuación se registra cada funcionalidad y cambio implementado durante la sesión, vinculándolo unívocamente a sus respectivos respaldos antes (`PRE`) y después (`POST`) del cambio:

### Hito 1: Infraestructura Base de Control de Roles (Fase 1 - Paso 0)
- **Descripción:** Implementación de funciones core de permisos en `modDatabase.bas`: `EsAdmin()`, `MostrarBloqueoPorRol()`, `ValidarPermisoAdmin()`. Eliminación de validaciones locales dispersas y centralización en la sesión de Access.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/modDatabase.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-17_2328_PRE_FASE1_PASO0.xlsm` (`F814E071502409B94A89C2FA545AC513`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-17_2330_POST_FASE1_PASO0.xlsm` (`DFDA9D433D53EC929518B561DA7CB7E5`)

### Hito 2: Desbloqueo Universal de Celdas (Fase 1 - Paso 1)
- **Descripción:** Eliminación de solicitud de contraseña manual (`InputBox`) en `MóduloDesbloqueoCeldas.bas`. Desbloqueo y rebloqueo automatizado mediante `ModuloConfigSegura.ObtenerClaveUniversalSegura()`.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/MóduloDesbloqueoCeldas.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-17_2344_PRE_FASE1_PASO1.xlsm` (`DFDA9D433D53EC929518B561DA7CB7E5`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-17_2346_POST_FASE1_PASO1.xlsm` (`BA1CB2218D2DCFD2AC284AA52E63E9CB`)

### Hito 3: Gestión de Contraseñas del Sistema (Fase 1 - Pasos 4 y 5)
- **Descripción:** Acceso directo para rol `ADMINISTRADOR` al panel de gestión de contraseñas de hojas y módulos. Bloqueo informativo para rol `OPERADOR`. Eliminación de `InputBox` en botones y scripts de prueba.
- **Archivos modificados:** `BotonGestionPasswords.bas`, `frmGestionPasswords.frm`, `ModuloPruebasPasswords.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-17_2351_PRE_FASE1_PASOS4Y5.xlsm` (`BA1CB2218D2DCFD2AC284AA52E63E9CB`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-17_2353_POST_FASE1_PASOS4Y5.xlsm` (`461D7CC14B8E33EAFEAE4FFB6221B8B5`)

### Hito 4: Criptografía y Configuración de Hashes (Fase 1 - Paso 6)
- **Descripción:** Verificación criptográfica SHA-256 en `ModuloConfPassHash.bas` condicionada a rol `ADMINISTRADOR`.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/ModuloConfPassHash.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-17_2358_PRE_FASE1_PASO6.xlsm` (`461D7CC14B8E33EAFEAE4FFB6221B8B5`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-17_2359_POST_FASE1_PASO6.xlsm` (`5CF3E19C04870183DE1B5DBA1CF736B7`)

### Hito 5: Borrado Masivo de Alojamientos (Fase 1 - Paso 7)
- **Descripción:** Protección de vaciado de tabla en `BorradoDatosAlojamiento.bas`. Bloqueo directo a rol `OPERADOR`, permitiendo la operación destructiva únicamente a `ADMINISTRADOR`.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0001_PRE_FASE1_PASO7.xlsm` (`5CF3E19C04870183DE1B5DBA1CF736B7`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0002_POST_FASE1_PASO7.xlsm` (`3B7621D574224386B6F5734889FC928B`)

### Hito 6: Panel de Control RGPD (Fase 1 - Paso 3)
- **Descripción:** Supresión de solicitud de contraseña manual para apertura de formulario RGPD. Comprobación directa de rol `ADMINISTRADOR` al pulsar botón RibbonX y en inicialización de `frmPanelRGPD`.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/ModuloRibbonRGPD.bas`, `frmPanelRGPD.frm`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0005_PRE_FASE1_PASO3.xlsm` (`3B7621D574224386B6F5734889FC928B`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0006_POST_FASE1_PASO3.xlsm` (`2A6C0958FF917F8B4F2BB617C4D04213`)

### Hito 7: Configuración de Búsqueda de Datos BDAS (Fase 1 - Paso 2)
- **Descripción:** Apertura directa del configurador `frmConfigBusquedaBDAS` para `ADMINISTRADOR` sin pedir contraseña adicional. Bloqueo por rol para `OPERADOR`.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/ActivarBusquedaDatosBDAS.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0012_PRE_FASE1_PASO2.xlsm` (`2A6C0958FF917F8B4F2BB617C4D04213`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0013_POST_FASE1_PASO2.xlsm` (`C6A062109B65ADB26F1BE77610D7E9C1`)

### Hito 8: Protección de Comentarios de Solicitudes (Fase 1 - Paso 10)
- **Descripción:** Corrección de vulnerabilidad de contraseña vacía en `MóduloComentarios.bas`. Asignación de contraseña segura predeterminada.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/MóduloComentarios.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0017_PRE_FASE1_PASO10.xlsm` (`C6A062109B65ADB26F1BE77610D7E9C1`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0018_POST_FASE1_PASO10.xlsm` (`0058CE54DB0B98999F3F2C1F24A0623A`)

### Hito 9: Saneamiento de Caracteres y Mojibake en `modDatabase.bas`
- **Descripción:** Limpieza exhaustiva de secuencias corruptas UTF-8 / Windows-1252 en la función de normalización de búsquedas y cadenas de texto.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/modDatabase.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0055_PRE_FIX_CARACTERES_MODDATABASE.xlsm` (`A40306B867E2A1A417B2ADB8A4D4C14F`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0058_POST_FIX_CARACTERES_MODDATABASE.xlsm` (`4282D6C258456AAB26D5E1A40989D0D4`)

### Hito 10: Saneamiento Cosmético Global de Cadenas de Texto
- **Descripción:** Corrección ortográfica y de tildes en módulos de calendarios, formato de fechas, lista negra y tarifas de unidad de descanso.
- **Archivos modificados:** Módulos de interfaz y utilidades.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0112_PRE_LIMPIEZA_COSMETICA.xlsm` (`4282D6C258456AAB26D5E1A40989D0D4`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0115_POST_LIMPIEZA_COSMETICA.xlsm` (`78EB55FBCB23046C81C0F9E91B384375`)

### Hito 11: Control de Visibilidad de Pestañas Nativas de Excel por Rol
- **Descripción:** Implementación en `customUI.xml` de `getVisible="RibbonNativeTabGetVisible"` para las 9 pestañas estándar de Excel (`TabHome`, `TabInsert`, `TabPageLayoutExcel`, `TabFormulas`, `TabData`, `TabReview`, `TabView`, `TabDeveloper`, `TabAddIns`).
  - **ADMINISTRADOR:** Las 9 pestañas nativas son visibles en tiempo de ejecución.
  - **OPERADOR:** Las 9 pestañas nativas permanecen ocultas, dejando disponibles únicamente las pestañas operativas de la aplicación.
  - **Botones de la aplicación:** Permanecen visibles y accesibles para ambos roles.
- **Archivos modificados:** `bdas-multiusuario/customUI/customUI.xml`, `bdas-multiusuario/vba-modules/modDatabase.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0205_PRE_PESTANAS_NATIVAS_ROLES.xlsm` (`1403BDC5AF52D42A2F68FD4677BD3F54`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0210_POST_PESTANAS_NATIVAS_ROLES.xlsm` (`B488080D0263C5FA7535E13DCF087692`)

### Hito 12: Protección Estricta de Hojas LOG por Rol
- **Descripción:** Cierre de vulnerabilidad que permitía a un usuario con rol `OPERADOR` eliminar filas o datos de las hojas de auditoría `LOG_GIJÓN`, `LOG_SOTO` y `LOG_OVIEDO`.
  - Todas las celdas de las hojas LOG configuradas como bloqueadas (`Cells.Locked = True`).
  - `AllowDeletingRows:=False, AllowDeletingColumns:=False, AllowInsertingRows:=False, AllowInsertingColumns:=False`.
  - Al iniciar sesión un `ADMINISTRADOR`, la hoja se desprotege para posibilitar tareas de mantenimiento; para `OPERADOR` se mantiene permanentemente blindada.
  - Gestión en eventos `Workbook_SheetActivate` y `Workbook_SheetDeactivate` en `ThisWorkbook.cls`.
- **Archivos modificados:** `BotonLOG.bas`, `ModuloLOG.bas`, `ThisWorkbook.cls`, `modDatabase.bas`, contenedor `BDAS_Multiusuario.xlsm`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0235_PRE_PROTECCION_HOJAS_LOG_ROLES.xlsm` (`37728133C4AD1740FC1D40B510E733F2`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0240_POST_PROTECCION_HOJAS_LOG_ROLES.xlsm` (`5017901D805E1876D02112A42457C0F2`)

### Hito 13: Purga Automática en Segundo Plano de Hojas LOG (>45 días)
- **Descripción:** Depuración automática y desasistida de registros con antigüedad superior a 45 días al abrir el libro en `Workbook_Open`.
  - Verificación inmediata en celda B2 (<1 ms): si el registro más antiguo no sobrepasa la retención, la rutina no itera.
  - Borrado en bloque continuo mediante `Rows("2:" & filaLimite).Delete`, minimizando el tiempo de procesamiento.
  - Ejecución en segundo plano sin parpadeos (`ScreenUpdating = False`).
- **Archivos modificados:** `ModuloLOG.bas`, `ThisWorkbook.cls`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0300_PRE_PURGA_AUTOMATICA_LOG.xlsm` (`5017901D805E1876D02112A42457C0F2`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0305_POST_PURGA_AUTOMATICA_LOG.xlsm` (`2B76736979710A428AF5F88705F3FF46`)

### Hito 14: Habilitación del Botón RibbonX "Último Nº Orden" en Hojas LOG
- **Descripción:** El botón de cinta RibbonX "Último Nº Orden" (`Button1` / `IrUltimaCeldaVacia`), al ser pulsado encontrándose en cualquiera de las hojas `LOG_GIJÓN`, `LOG_SOTO` o `LOG_OVIEDO`, posiciona el cursor en la última celda con datos de la columna A y enfoca verticalmente la pantalla (`Application.Goto ..., True`).
- **Archivos modificados:** `bdas-multiusuario/vba-modules/BotonIrUltimaCeldaOrden.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_0310_PRE_IR_ULTIMA_CELDA_LOG.xlsm` (`2B76736979710A428AF5F88705F3FF46`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_0315_POST_IR_ULTIMA_CELDA_LOG.xlsm` (`2553D789AB30EE062203062265F37712`)

### Hito 15: Formulario de Gestión de Lista Negra y Sincronización Automática con Hoja Protegida y Access
- **Descripción:** Implementación de la solución arquitectónica integral para la Lista Negra:
  1. Creación del nuevo UserForm interactivo [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm) con visualización en `ListBox`, alta/actualización y retirada de vetos (baja lógica `Activo = False`).
  2. Sincronización automática de la hoja `LISTA NEGRA` al abrir y operar desde la base central Access `Residencia_BE.accdb`.
  3. Modificación del botón de cinta en [BotonListaNegra.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonListaNegra.bas) para sincronizar hoja y abrir el formulario modal.
  4. Saneamiento de registros residuales y normalización de fechas históricas en la tabla `ListaNegra` de Access.
- **Archivos modificados:** `frmGestionListaNegra.frm`, `BotonListaNegra.bas`, `modDatabase.bas`, base Access `Residencia_BE.accdb`.
- **Backup PRE XLSM:** `BDAS_Multiusuario_BACKUP_2026-10-18_1937_PRE_GESTION_LISTA_NEGRA.xlsm` (`2553D789AB30EE062203062265F37712`)
- **Backup PRE ACCDB:** `Residencia_BE_BACKUP_2026-10-18_1937_PRE_GESTION_LISTA_NEGRA.accdb` (`ADDBB2E73CAB45CAF06F3F830DB08840`)
- **Backup POST XLSM:** `BDAS_Multiusuario_BACKUP_2026-10-18_1945_POST_GESTION_LISTA_NEGRA.xlsm` (`38D2609E98AC6FEA5F7387BB8A5DDF99`)
- **Backup POST ACCDB:** `Residencia_BE_BACKUP_2026-10-18_1945_POST_GESTION_LISTA_NEGRA.accdb` (`9C6CCB9E3819CF8425989596436F42AC`)

### Hito 16: Corrección de Ámbito en Función Auxiliar `modDatabase.Nz` (Public)
- **Descripción:** Resolución de error de compilación COM en tiempo de carga modal (*"No se encontró el método o el dato miembro"*) cambiando la visibilidad de `Private Function Nz` a `Public Function Nz` en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas), haciéndola accesible para el formulario y futuros módulos.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/modDatabase.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_1958_PRE_FIX_NZ_LISTA_NEGRA.xlsm` (`38D2609E98AC6FEA5F7387BB8A5DDF99`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2000_POST_FIX_NZ_LISTA_NEGRA.xlsm` (`AAC33447321A0F79DE91FC7CDF3233C5`)

### Hito 17: Escucha de Eventos COM `WithEvents`, Botones "Buscar" y "Limpiar", y Corrección Ortográfica
- **Descripción:** 
  1. Enlace de eventos de controles dinámicos en tiempo de ejecución mediante variables con ámbito de módulo `WithEvents` (`btnAnadir`, `btnRetirar`, `btnCerrar`, `btnBuscar`, `btnLimpiar`, `txtBuscar`, `lstLista`).
  2. Adición física de botones dedicados **Buscar** y **Limpiar** junto al cuadro de búsqueda y contador dinámico.
  3. Soporte para disparo de búsqueda al presionar la tecla `Enter` (`KeyCode = 13`).
  4. Corrección ortográfica estricta sustituyendo "Anadir" por "Añadir".
- **Archivos modificados:** `bdas-multiusuario/vba-modules/frmGestionListaNegra.frm`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2015_PRE_FIX_EVENTOS_LISTA_NEGRA.xlsm` (`AAC33447321A0F79DE91FC7CDF3233C5`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2018_POST_FIX_EVENTOS_LISTA_NEGRA.xlsm` (`F4FE1BC1B1674D32B34D8FE944A955A6`)

### Hito 18: Optimización Extrema del Rendimiento en Lista Negra (Bulk Array Write)
- **Descripción:** Supresión de cuellos de botella de latencia COM (>9 segundos) al pulsar el botón Lista Negra:
  1. Sustitución de escritura celda a celda por volcado masivo en memoria matriz Variant bidimensional (`GetRows` -> `arrData`) y volcado en bloque único `ws.Range("A3").Resize(numReg, 3).Value = arrData`.
  2. Restricción del bloqueo de celdas estrictamente al rango ocupado (`A1:C & (2+numReg)`).
  3. Desactivación de refresco de pantalla y recálculo automático (`ScreenUpdating = False`, `EnableEvents = False`).
  4. Poblado directo de `ListBox` mediante asignación de matriz en `frmGestionListaNegra.frm`.
  5. **Resultado cuantitativo:** Tiempo de sincronización reducido de **9,031 s a 0,289 s** (Aceleración de **31x / >3.100%**).
- **Archivos modificados:** `modDatabase.bas`, `frmGestionListaNegra.frm`, `BotonListaNegra.bas`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2032_PRE_OPTIMIZACION_LISTA_NEGRA.xlsm` (`F4FE1BC1B1674D32B34D8FE944A955A6`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2041_POST_OPTIMIZACION_LISTA_NEGRA.xlsm` (`AEF8A396CB135CFD41A0EE0803406FE0`)

### Hito 19: Preselección Total de Residencias Asignadas por Defecto en el Selector
- **Descripción:** Modificación de `frmSelectorResidencia.frm` en `UserForm_Initialize` para iterar y marcar automáticamente con tick (`Selected(i) = True`) todas las residencias autorizadas al usuario, permitiendo acceder de inmediato con `Enter` sin requerir marcación manual individual.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/frmSelectorResidencia.frm`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2105_PRE_SELECTOR_TODAS_DEFAULT.xlsm` (`AEF8A396CB135CFD41A0EE0803406FE0`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2108_POST_SELECTOR_TODAS_DEFAULT.xlsm` (`6E05177A99727A64C3937594DBA7F0E2`)

### Hito 20: Navegación 100% por Teclado (Enter / Tab) en Formulario de Login
- **Descripción:** Implementación de ciclo continuo de teclado en `frmLogin.frm`:
  1. Posicionamiento automático del foco en campo Usuario al iniciar (`txtUsuario.SetFocus`).
  2. Avance secuencial con `Enter` (código 13) o `Tab`: Usuario ➔ Contraseña ➔ Botón Entrar.
  3. Ejecución directa de la validación al pulsar `Enter` sobre el botón Entrar.
  4. Preservación absoluta del funcionamiento con ratón y compatibilidad total sin efectos colaterales.
- **Archivos modificados:** `bdas-multiusuario/vba-modules/frmLogin.frm`.
- **Backup PRE:** `BDAS_Multiusuario_BACKUP_2026-10-18_2118_PRE_LOGIN_KEYBOARD_NAV.xlsm` (`6E05177A99727A64C3937594DBA7F0E2`)
- **Backup POST:** `BDAS_Multiusuario_BACKUP_2026-10-18_2121_POST_LOGIN_KEYBOARD_NAV.xlsm` (`4D3D5F3C2120B9E0F91D1CD2F05B9945`)

### Hito 21: Consolidación y Certificación Oficial de la Versión de Referencia (Cierre 21:25)
- **Descripción:** Elevación del estado plenamente probado, optimizado y sin errores a nueva **Versión de Referencia Oficial** del proyecto:
  1. Actualización de réplicas maestras en `H:\ResidenciaBD\Plantilla\VersionReferencia\` en modo Solo Lectura estricto (`IsReadOnly = True`).
  2. Archivamiento de la versión de referencia previa (Hito 03:20) en el histórico `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`.
  3. Generación de copia de seguridad consolidada de cierre de jornada en `H:\ResidenciaBD\BackupsDiarios\`.
  4. Actualización del script automatizado de restauración [Restaurar-Version-Referencia.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Version-Referencia.ps1).
- **Archivos modificados:** Réplicas protegidas de referencia y scripts de gobierno.
- **Backup Consolidado XLSM:** `BDAS_Multiusuario_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.xlsm` (`4D3D5F3C2120B9E0F91D1CD2F05B9945`)
- **Backup Consolidado ACCDB:** `Residencia_BE_BACKUP_2026-10-18_2125_VERSION_REFERENCIA.accdb` (`2AB9DAEB5015BE53738F23743D309228`)

---

## 3. Matriz Resumen de Trazabilidad Criptográfica Completa (2026-10-18)

A continuación se presenta la tabla exhaustiva que vincula cada una de las copias de seguridad incrementales realizadas durante el día con su respectiva intervención y huella criptográfica verificada:

| Nº | Funcionalidad / Cambio Implementado | Archivo de Backup PRE | Hash MD5 PRE | Archivo de Backup POST | Hash MD5 POST |
|:---:|:---|:---|:---:|:---|:---:|
| **1** | Infraestructura Roles (Paso 0) | `..._2328_PRE_FASE1_PASO0.xlsm` | `F814E071502409B94A89C2FA545AC513` | `..._2330_POST_FASE1_PASO0.xlsm` | `DFDA9D433D53EC929518B561DA7CB7E5` |
| **2** | Desbloqueo Universal (Paso 1) | `..._2344_PRE_FASE1_PASO1.xlsm` | `DFDA9D433D53EC929518B561DA7CB7E5` | `..._2346_POST_FASE1_PASO1.xlsm` | `BA1CB2218D2DCFD2AC284AA52E63E9CB` |
| **3** | Gestión Contraseñas (Pasos 4 y 5) | `..._2351_PRE_FASE1_PASOS4Y5.xlsm` | `BA1CB2218D2DCFD2AC284AA52E63E9CB` | `..._2353_POST_FASE1_PASOS4Y5.xlsm` | `461D7CC14B8E33EAFEAE4FFB6221B8B5` |
| **4** | Hashes SHA256 (Paso 6) | `..._2358_PRE_FASE1_PASO6.xlsm` | `461D7CC14B8E33EAFEAE4FFB6221B8B5` | `..._2359_POST_FASE1_PASO6.xlsm` | `5CF3E19C04870183DE1B5DBA1CF736B7` |
| **5** | Borrado Total Tabla (Paso 7) | `..._0001_PRE_FASE1_PASO7.xlsm` | `5CF3E19C04870183DE1B5DBA1CF736B7` | `..._0002_POST_FASE1_PASO7.xlsm` | `3B7621D574224386B6F5734889FC928B` |
| **6** | Panel RGPD (Paso 3) | `..._0005_PRE_FASE1_PASO3.xlsm` | `3B7621D574224386B6F5734889FC928B` | `..._0006_POST_FASE1_PASO3.xlsm` | `2A6C0958FF917F8B4F2BB617C4D04213` |
| **7** | Búsqueda BDAS (Paso 2) | `..._0012_PRE_FASE1_PASO2.xlsm` | `2A6C0958FF917F8B4F2BB617C4D04213` | `..._0013_POST_FASE1_PASO2.xlsm` | `C6A062109B65ADB26F1BE77610D7E9C1` |
| **8** | Protección Comentarios (Paso 10) | `..._0017_PRE_FASE1_PASO10.xlsm` | `C6A062109B65ADB26F1BE77610D7E9C1` | `..._0018_POST_FASE1_PASO10.xlsm` | `0058CE54DB0B98999F3F2C1F24A0623A` |
| **9** | Saneamiento `modDatabase.bas` | `..._0055_PRE_FIX_CARACTERES_...xlsm` | `A40306B867E2A1A417B2ADB8A4D4C14F` | `..._0058_POST_FIX_CARACTERES_...xlsm` | `4282D6C258456AAB26D5E1A40989D0D4` |
| **10** | Saneamiento Cosmético Global | `..._0112_PRE_LIMPIEZA_COSMETICA.xlsm` | `4282D6C258456AAB26D5E1A40989D0D4` | `..._0115_POST_LIMPIEZA_COSMETICA.xlsm` | `78EB55FBCB23046C81C0F9E91B384375` |
| **11** | Pestañas Nativas por Rol | `..._0205_PRE_PESTANAS_NATIVAS...xlsm` | `1403BDC5AF52D42A2F68FD4677BD3F54` | `..._0210_POST_PESTANAS_NATIVAS...xlsm` | `B488080D0263C5FA7535E13DCF087692` |
| **12** | Protección Hojas LOG por Rol | `..._0235_PRE_PROTECCION_HOJAS...xlsm` | `37728133C4AD1740FC1D40B510E733F2` | `..._0240_POST_PROTECCION_HOJAS...xlsm` | `5017901D805E1876D02112A42457C0F2` |
| **13** | Purga Automática LOG (>45 días) | `..._0300_PRE_PURGA_AUTOMATICA...xlsm` | `5017901D805E1876D02112A42457C0F2` | `..._0305_POST_PURGA_AUTOMATICA...xlsm` | `2B76736979710A428AF5F88705F3FF46` |
| **14** | Botón "Último Nº Orden" en LOG | `..._0310_PRE_IR_ULTIMA_CELDA...xlsm` | `2B76736979710A428AF5F88705F3FF46` | `..._0315_POST_IR_ULTIMA_CELDA...xlsm` | `2553D789AB30EE062203062265F37712` |
| **15** | Formulario Gestión Lista Negra | `..._1937_PRE_GESTION_LISTA_NEGRA.xlsm` | `2553D789AB30EE062203062265F37712` | `..._1945_POST_GESTION_LISTA_NEGRA.xlsm` | `38D2609E98AC6FEA5F7387BB8A5DDF99` |
| — | *Base Access Lista Negra* | `Residencia_BE...1937_PRE...accdb` | `ADDBB2E73CAB45CAF06F3F830DB08840` | `Residencia_BE...1945_POST...accdb` | `9C6CCB9E3819CF8425989596436F42AC` |
| **16** | Corrección `modDatabase.Nz` (Public) | `..._1958_PRE_FIX_NZ_LISTA_NEGRA.xlsm` | `38D2609E98AC6FEA5F7387BB8A5DDF99` | `..._2000_POST_FIX_NZ_LISTA_NEGRA.xlsm` | `AAC33447321A0F79DE91FC7CDF3233C5` |
| **17** | Eventos COM y Botón Buscar / Añadir | `..._2015_PRE_FIX_EVENTOS_...xlsm` | `AAC33447321A0F79DE91FC7CDF3233C5` | `..._2018_POST_FIX_EVENTOS_...xlsm` | `F4FE1BC1B1674D32B34D8FE944A955A6` |
| **18** | Optimización 31x Matriz Memoria | `..._2032_PRE_OPTIMIZACION_...xlsm` | `F4FE1BC1B1674D32B34D8FE944A955A6` | `..._2041_POST_OPTIMIZACION_...xlsm` | `AEF8A396CB135CFD41A0EE0803406FE0` |
| **19** | Selector Preselección por Defecto | `..._2105_PRE_SELECTOR_TODAS_...xlsm` | `AEF8A396CB135CFD41A0EE0803406FE0` | `..._2108_POST_SELECTOR_TODAS_...xlsm` | `6E05177A99727A64C3937594DBA7F0E2` |
| **20** | Navegación Teclado Login (Enter/Tab) | `..._2118_PRE_LOGIN_KEYBOARD_...xlsm` | `6E05177A99727A64C3937594DBA7F0E2` | `..._2121_POST_LOGIN_KEYBOARD_...xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` |
| **21** | **Versión de Referencia Oficial** | *(Estado Final Consolidado)* | — | `..._2125_VERSION_REFERENCIA.xlsm` | `4D3D5F3C2120B9E0F91D1CD2F05B9945` |
| — | *Base Access Referencia Final* | *(Estado Final Consolidado)* | — | `Residencia_BE...2125_VERSION_REF.accdb` | `2AB9DAEB5015BE53738F23743D309228` |

---

## 4. Estado de Cumplimiento de Reglas del Proyecto

- **Regla estricta de nombres:** Cumplimiento total. Se utiliza el texto literal `GIJÓN` (`\xd3` / carácter 211) en todo el código fuente y nombres de hojas (`LOG_GIJÓN`, `RESIDENCIA GIJÓN`), prohibiéndose de forma estricta cualquier concatenación tipo `Chr(211)`.
- **Inmutabilidad del Golden Master:** La réplica Golden Master (`H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm`) ha permanecido intacta y sin modificaciones (MD5: `A03DB5430C8E913AB2703F5116904DD2`).
- **Respaldo de Referencia Previo:** Preservado y archivado en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_0320.xlsm`.
- **Integridad COM y Compilación VBA:** Validada en Excel 2016 con **`0 ERRORES`**.

