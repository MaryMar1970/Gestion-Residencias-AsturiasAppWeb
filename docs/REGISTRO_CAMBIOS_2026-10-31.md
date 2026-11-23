# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-10-31 | Hito Oficial: Versión de Referencia 2026-10-31 20:15
### Ámbito: Control de Concurrencia Multiusuario Completo (V5 + V5.1 + V5.2 Cierre Definitivo) + Blindaje y Homogeneización Integral de Calendarios (GIJÓN, SOTO, OVIEDO)

---

## 1. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia exacta entre cada hito técnico de implementación, los archivos de respaldo incrementales generados en disco, sus dimensiones, hashes criptográficos y el conjunto de funcionalidades y correcciones introducidas:

| Hito / Copia Incremental | Ruta del Archivo de Respaldo | Tamaño (Bytes) | Hash MD5 / SHA-256 | Funcionalidades, Operaciones y Módulos Vinculados |
| :--- | :--- | :--- | :--- | :--- |
| **0. Estado Base de Partida**<br>*(Hito 2026-10-29 20:45)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-29_2045.xlsm` | 6.796.234 B | **MD5:** `D203D5058F41ED79FC3648C7309D8B51`<br>**SHA256:** `6E3C5C0368E2D5DEFB56B8046CC456D54C26F77B2E6F5C0FE47669BE982231AD` | Punto de partida previo a la implementación del control de concurrencia avanzado. Incorporaba sincronización en tiempo real de 16 operaciones y herramientas RibbonX V4. |
| **1. Backup Preventivo PRE-V5**<br>*(Antes de Opción B Concurrencia)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-30_PRE_CONCURRENCIA_V5.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-30_PRE_CONCURRENCIA_V5.xlsm` | 6.782.775 B | **MD5:** `4CACC3B96234101272D811D14C58F64E`<br>**SHA256:** `6B227AFD668AA51EE5D2DB8DFE542C01EAF33A7BD2EE9A3B2153216BBF240466` | Instantánea preventiva de seguridad previa al inicio de las operaciones quirúrgicas `OP-CONC-01` a `OP-CONC-07`. |
| **2. Backup POST-V5**<br>*(Concurrencia Opción B Consolidada)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-30_POST_CONCURRENCIA_V5.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-30_POST_CONCURRENCIA_V5.xlsm` | 6.831.066 B | **MD5:** `95DF2EAF979679271DA91F60EE6374BD`<br>**SHA256:** `59AAC754A21184F04BAAB06C429A5749B8FBE940CCB4231E746ACC8DCBE61F3A` | **Implementación del Control de Concurrencia Opción B:**<br>1. `OP-CONC-01`: Inyección de constantes de Lease Lock (120 seg) y 7 funciones en `modDatabase.bas` (`AsegurarEsquemaConcurrenciaBD`, `EstaExpedienteBloqueadoPorOtro`, `IntentarReservarOrdenBD`, `ReservarSiguienteOrdenLibreBD`, `GestionarEntradaNumOrdenHoja`, `AsegurarOrdenExiste`, `SincronizarFilaHaciaAccess`).<br>2. `OP-CONC-02`: Actualización de firma y predicados en `EliminarOrdenBD`.<br>3. `OP-CONC-03`: Trazabilidad de operador en `InsertarLog`.<br>4. `OP-CONC-04`: Funciones de operador en `ModuloSeguridadLocal.bas`.<br>5. `OP-CONC-05`, `06`, `07`: Delegación del evento `Case 1` en `GestionarEntradaNumOrdenHoja` y blindaje contra edición múltiple con `Application.Undo` en `Hoja2` (GIJÓN), `Hoja20` (SOTO) y `Hoja24` (OVIEDO). |
| **3. Backup Preventivo PRE-V5.1**<br>*(Antes de Correcciones V5.1)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_PRE_CONCURRENCIA_V5_1.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-31_PRE_CONCURRENCIA_V5_1.xlsm` | 6.716.865 B | **MD5:** `6B99B0D1F484FFC00B80301F24F627EF`<br>**SHA256:** `B777CAFDDF578D0232815FC6925F124E67DCF0CE434353AD953F8F82FB8D589B` | Estado inicial tras apertura matinal del proyecto antes de acometer la fase de ajuste de alineación de esquemas y robustez de llamadas. |
| **4. Backup POST-V5.1**<br>*(Ajuste de Esquema y Robustez)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CONCURRENCIA_V5_1.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CONCURRENCIA_V5_1.xlsm` | 6.736.083 B | **MD5:** `9974EC62558E6DAA68A433E0DD77CDC6`<br>**SHA256:** `9EE282DDBB67FDAD7BB1637D1BFF91AE60EE2D0EBBB3AFA443874699BDFA9ED1` | **Correcciones de Cierre V5.1:**<br>1. `OP-C51-01`: Detección preventiva de columnas `UsuarioModificacion` y `UltimaModificacion` en Access con autocreación DDL.<br>2. `OP-C51-02`: Inclusión de `usuarioActual` y parámetros de lease en `EstaExpedienteBloqueadoPorOtro`.<br>3. `OP-C51-03`: Normalización del diálogo de adjudicación de orden alternativa.<br>4. `OP-C51-04` y `05`: Ajustes de tipos en hojas de residencia. |
| **5. Backup Preventivo PRE-V5.2**<br>*(Antes de Cierre Definitivo)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_PRE_CONCURRENCIA_V5_2.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-31_PRE_CONCURRENCIA_V5_2.xlsm` | 6.736.083 B | **MD5:** `9974EC62558E6DAA68A433E0DD77CDC6`<br>**SHA256:** `9EE282DDBB67FDAD7BB1637D1BFF91AE60EE2D0EBBB3AFA443874699BDFA9ED1` | Instantánea de partida antes de aplicar el paquete de cierre definitivo V5.2. |
| **6. Backup POST-V5.2**<br>*(Cierre Definitivo Concurrencia)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CONCURRENCIA_V5_2.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CONCURRENCIA_V5_2.xlsm` | 6.755.017 B | **MD5:** `9CF17882371F25E816DD78F5D0001332`<br>**SHA256:** `6FD5A725F9C5B300C7733A88F3794EF5F67E5424BB8FAB7DBE2BC811D5653598` | **Cierre Definitivo V5.2 (Token GUID + Predicados Atómicos + Detección Dinámica):**<br>1. `OP-C52-01A`: API Win32 `CoCreateGuid` en `ole32.dll` + tipo `GUID_TYPE`.<br>2. `OP-C52-01B`: Token de sesión de instancia `Usuario@NombrePC#8HexGUID` en `ObtenerTokenSesionInstancia()`.<br>3. `OP-C52-01C`: Formateador robusto de propietario de bloqueo `FormatearPropietarioBloqueo()`.<br>4. `OP-C52-02`: Predicado atómico en `EliminarOrdenBD` contra bloqueos activos de terceros.<br>5. `OP-C52-03`: Mapeo exhaustivo de 28 columnas en `SincronizarFilaHaciaAccess` con comprobación de Lease Lock.<br>6. `OP-C52-04`: `LimpiarCacheTokenSesion` en `ModuloSeguridadLocal.bas`.<br>7. `OP-C52-05`: Detección reactiva de límites reales de columna (`ultimaColReal`) en `Hoja2`, `Hoja20`, `Hoja24`. |
| **7. Backup Intermedio 18:30**<br>*(Referencia Archivada)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_1830.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_1830_VERSION_REFERENCIA.xlsm` | 6.755.017 B | **MD5:** `9CF17882371F25E816DD78F5D0001332`<br>**SHA256:** `6FD5A725F9C5B300C7733A88F3794EF5F67E5424BB8FAB7DBE2BC811D5653598` | Hito intermedio de consolidación del Control de Concurrencia archivado en Histórico de Versiones de Referencia. |
| **8. Backup POST-CALENDARIOS-FIX**<br>*(Corrección y Homogeneización)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CALENDARIOS_FIX.xlsm`<br>`H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario_BACKUP_2026-10-31_POST_CALENDARIOS_FIX.xlsm` | 6.749.596 B | **MD5:** `6AE80126541D7D7F273B94929ECAE6A9`<br>**SHA256:** `210745C9D089798C1623A4CF0E448FAA7F4344C41A7B6A7B0AAA6B6B7F378E9A` | **Solución Definitiva de Protección y Homogeneización de Calendarios:**<br>1. Erradicación del error en tiempo de ejecución 1004 (*«La contraseña no es válida...»*) al activar `Calendario GIJÓN`.<br>2. Desprotección segura mediante `ModuloConfigSegura.ObtenerPasswordHojas()` en `ActualizarCalendarioInterno` para Gijón, Soto y Oviedo.<br>3. Reprotección incondicional robusta con `UserInterfaceOnly:=True, AllowFormattingCells:=True`.<br>4. Reparación colateral de fallos en `ModuloGestionHabitaciones` y `Hoja41` (Festivos). |
| **9. Versión de Referencia Oficial Vigente**<br>*(Hito 2026-10-31 20:15)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-31_2015.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-31_2015_VERSION_REFERENCIA.xlsm` | 6.749.596 B | **MD5:** `6AE80126541D7D7F273B94929ECAE6A9`<br>**SHA256:** `210745C9D089798C1623A4CF0E448FAA7F4344C41A7B6A7B0AAA6B6B7F378E9A` | **Versión de Referencia Oficial Activa en Producción.** Protegida en solo lectura, replicada en red (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`) y vinculada con `Residencia_BE.accdb` (6.746.112 B \| MD5: `6237949FA8E9F503C9623F0A7294310D`). |

---

## 2. Detalle Exhaustivo de las Modificaciones por Módulo

### 2.1. `modDatabase.bas`

#### A. Identificador Único de Sesión por Instancia (Win32 GUID)
- **Problema previo:** Si el mismo operador abría la aplicación en dos puestos diferentes (o si un usuario compartía credencial de red), el sistema interpretaba que eran el "mismo operador" y permitía que una sesión pisara las modificaciones o bloqueos de la otra.
- **Solución implementada:**
  - Declaración nativa de la API Win32:
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
  - Función `ObtenerTokenSesionInstancia()`: Genera una cadena compacta `Usuario@NombrePC#8HexGUID` (≤ 50 caracteres para adecuarse al campo `UsuarioModificacion TEXT(50)` de Access). Cada libro abierto genera su propio token irrepetible.

#### B. Formateador Robusto de Propietario de Bloqueo
- **Implementación:** Función `FormatearPropietarioBloqueo(ByVal rawToken As String) As String`:
  Separa de forma limpia el nombre del operador, el nombre del PC y el token para mostrar en las ventanas emergentes un mensaje claro y comprensible (ejemplo: `"El operador admin (en PC-SALA-02) está modificando este expediente..."`), evitando que el usuario vea identificadores hexadecimales crudos.

#### C. Predicados Atómicos en `EliminarOrdenBD` y `SincronizarFilaHaciaAccess`
- **Eliminación Segura:**
  Antes de ejecutar el archivado y borrado de una orden en Access, se evalúa si el registro está bloqueado activamente por otra sesión:
  ```vba
  Dim tokenActual As String
  tokenActual = ObtenerTokenSesionInstancia()
  If EstaExpedienteBloqueadoPorOtro(rs, tokenActual) Then
      MsgBox "No se puede eliminar la Orden " & numOrden & " porque está siendo editada por " & _
             FormatearPropietarioBloqueo(rs("UsuarioModificacion").Value) & "...", vbExclamation, "Operación Bloqueada"
      Exit Function
  End If
  ```
- **Sincronización Indexada de 28 Columnas:**
  En `SincronizarFilaHaciaAccess`:
  - Se verifica `EstaExpedienteBloqueadoPorOtro`.
  - Se distingue entre registro preexistente (`esRegistroExistente = True`) y nueva inserción (`rs.AddNew`).
  - Se asignan de manera segura y tipada las 28 columnas de la tabla `Ordenes` (DNI, Nombre, Apellidos, Tipo, Teléfono, Vehículo, Matrícula, Fechas, Tarifas, Importes, Facturación, etc.).
  - Se actualizan atómicamente `UsuarioModificacion = tokenActual` y `UltimaModificacion = Now()`.

---

### 2.2. `ModuloSeguridadLocal.bas`

- Se incluyó la rutina de reinicio de sesión `LimpiarCacheTokenSesion()` para permitir que los cambios de sesión o recargas del sistema invaliden el token en memoria sin reiniciar la aplicación Excel.
- Vinculación de `ObtenerUsuarioActual()` como base del token de sesión generado en `modDatabase.bas`.

---

### 2.3. Hojas Operativas de Residencia (`Hoja2` GIJÓN, `Hoja20` SOTO, `Hoja24` OVIEDO)

#### A. Detección Dinámica de Dimensiones Reales en Filas
- **Problema previo:** En resoluciones o formatos variables, la lectura de columnas asumía arrays de tamaño prefijado, provocando ocasionalmente desbordamientos o desalineación en celdas vacías.
- **Solución implementada:**
  - Sustitución por cálculo reactivo de límites reales:
    ```vba
    Dim ultimaColReal As Long
    ultimaColReal = Me.Cells(targetRow, Me.Columns.Count).End(xlToLeft).Column
    ```
  - Iteración segura columna a columna respetando la estructura real de la hoja sin saltos ni lecturas truncadas.

#### B. Reserva Provisional Offline con Recuperación Reactiva
- Si una estación de trabajo pierde momentáneamente la conectividad con el servidor Access (`Residencia_BE.accdb`), la función de gestión de orden no bloquea la interfaz de usuario: adjudica localmente la reserva provisional y programa la confirmación atómica en la base de datos en cuanto se reanuda el enlace de red.

#### C. Blindaje Antipegado y Edición Masiva en Columna A
- Intercepción en `Case 1`:
  ```vba
  ElseIf Target.Cells.Count > 1 Then
      On Error Resume Next
      Application.Undo
      On Error GoTo ErrHandler_Change
      MsgBox "Por seguridad multiusuario, no se permite pegar ni modificar múltiples Nº ORDEN a la vez en la Columna A.", _
             vbExclamation, "Operación Bloqueada"
  End If
  ```
  Garantiza que ningún arrastre accidental, pegado desde portapapeles o borrado de rango destruya la unicidad ni desincronice el backend.

---

### 2.4. Módulos de Calendario (`ModuloCalendarioGijon`, `ModuloCalendarioSoto`, `ModuloCalendarioOviedo`)

#### A. Causa Raíz Identificada
Al iniciar la aplicación con usuario `ADMINISTRADOR` seleccionando todas las residencias, el flujo de navegación aterriza por defecto en la hoja `Calendario GIJÓN`. Su evento `Worksheet_Activate` invoca `ActualizarCalendarioFuturoGijon`, que delega en `ActualizarCalendarioInterno`.
En dicha rutina:
1. Se utilizaba una desprotección con contraseña en texto plano no alineada con el almacén seguro centralizado, desencadenando el **Error 1004 en tiempo de ejecución: «La contraseña especificada no es válida»**.
2. Al fallar la desprotección, cualquier intento de limpiar celdas (`wsCal.Range(...).ClearContents`) o repintar ocupaciones en celdas bloqueadas provocaba la caída del procedimiento.
3. Existía disparidad de comportamiento respecto a las hojas de Soto y Oviedo.

#### B. Solución Quirúrgica y Homogénea Aplicada
En los procedimientos `ActualizarCalendarioInterno` de los tres módulos se implementó:
1. **Recuperación Oficial de Contraseña:**
   ```vba
   pwdHojas = ModuloConfigSegura.ObtenerPasswordHojas()
   ```
2. **Desprotección Defensiva en Cascada:**
   ```vba
   estabaProtegida = wsCal.ProtectContents
   If Len(pwdHojas) > 0 Then wsCal.Unprotect password:=pwdHojas
   If wsCal.ProtectContents Then wsCal.Unprotect password:=""
   If wsCal.ProtectContents Then wsCal.Unprotect
   ```
3. **Reprotección Incondicional con Modo Macros (`UserInterfaceOnly:=True`):**
   Tanto en la salida limpia (`CleanExit`) como en el controlador de errores (`ErrorHandler`):
   ```vba
   If Not wsCal Is Nothing Then
       wsCal.Protect password:=pwdHojas, UserInterfaceOnly:=True, AllowFormattingCells:=True
   End If
   ```

#### C. Beneficios Clave
- **Cero Errores 1004:** Se erradica por completo la ventana de error al entrar a Gijón, Soto u Oviedo.
- **Protección para Usuarios, Vía Libre para Macros:** El usuario final no puede editar manualmente las celdas del calendario (evitando corrupción de cuadrícula), pero el motor VBA puede escribir, repintar y formatear en milisegundos sin solicitar contraseñas.
- **Reparación de Efectos Secundarios Latentes:** Se resuelven de forma preventiva los bloqueos de ejecución en `ModuloGestionHabitaciones` (al ocultar/mostrar habitaciones) y en `Hoja41` (al actualizar días festivos).
- **Simetría Total:** Las tres residencias operan bajo el mismo estándar de robustez.

---

## 3. Pruebas de Calidad, Integridad y Gobernanza Técnica

1. **Compilación COM de VBA en Excel:**
   - Comando ejecutado: `xl.VBE.CommandBars.FindControl(1, 578).Enabled`
   - **Resultado:** `False` (0 errores de sintaxis, tipado o dependencias en los 173 componentes).
2. **Control Anti-Mojibake y Juego de Caracteres:**
   - Escaneo integral byte a byte en Windows-1252 ANSI sin BOM.
   - **Resultado:** 0 caracteres corruptos introducidos en los módulos operativos.
3. **Regla Estricta de Nombres (Toponimia Asturiana):**
   - Búsqueda exhaustiva del término `GIJÓN`.
   - **Resultado:** 100% de uso del texto literal `GIJÓN`. 0 ocurrencias de `Chr(211)` o concatenaciones anómalas.
4. **Protección de Estructura e Inmutabilidad de Referencias:**
   - La réplica Golden Master (`H:\ResidenciaBD\Plantilla\GoldenMaster\`) se mantiene 100% intacta e inmutable.
   - La nueva Versión de Referencia Oficial (`H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`) ha quedado protegida con el atributo de **Solo Lectura** (`IsReadOnly = True`).
   - La copia de red para usuarios concurrentes (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`) ha quedado sincronizada y accesible.
