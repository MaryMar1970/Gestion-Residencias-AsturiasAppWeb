# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-11-03 | Hito Oficial Consolidado: Versión de Referencia 2026-11-03 00:15
### Ámbito: Resolución Dinámica Backend DB_PATH + Idempotencia EliminarOrdenBD y Saneamiento WAL SICO-N2 + Protección Hojas con UserInterfaceOnly en FechasPeticion y ListasDinamicasCPyPoblaciones + Coherencia Habitacional GIJÓN/OVIEDO

---

## 1. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia exacta entre cada hito técnico de implementación en la jornada 2026-11-02 (cerrada a las 00:15 del 2026-11-03), los archivos de respaldo incrementales generados en disco, sus dimensiones, hashes criptográficos y el conjunto de funcionalidades y correcciones vinculadas:

| Hito / Copia Incremental | Ruta del Archivo de Respaldo | Tamaño (Bytes) | Hashes Criptográficos (MD5 / SHA-256) | Funcionalidades, Operaciones y Módulos Vinculados |
| :--- | :--- | :--- | :--- | :--- |
| **0. Estado Base de Partida**<br>*(Hito 2026-11-02 03:30)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-11-02_0330.accdb` | 6.430.806 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `F73F5C628319611FD8C56CA773E60B67`<br>**XLSM SHA256:** `749A06FF1C80B1664E5CC8B22CDC00BDBA1EA3D4E96BDB0BE033E31557787A60`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Punto de partida de la sesión. Incorporaba Cierre Definitivo SICO-N2, Anti-Bloqueo DWM, Gestión Borrado RGPD, Reintegración de Oviedo y Blindaje de Gestión de Usuarios en `modDatabase.bas` (Error 3704). |
| **1. Backup Preventivo PRE-RUTAS**<br>*(Antes de Rutas Dinámicas)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2202_PRE_RUTAS.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2202_PRE_RUTAS.accdb` | 6.379.116 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `0C99D3A989A4F7B22B8255CCB575EBB8`<br>**XLSM SHA256:** `21C4E9F95ABD46FA1941C9C4EF36A00314C24E11A9ADD2D98DEF6F90111E85B`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Copia de seguridad preventiva antes de desacoplar las constantes fijas de red `H:\GIJ1\...` de la base de datos backend. |
| **2. Backup POST-RUTAS**<br>*(Rutas Dinámicas Implementadas)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2202_POST_RUTAS.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2202_POST_RUTAS.accdb` | 6.385.789 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `549921707952BC6EAA423A6CBC107500`<br>**XLSM SHA256:** `9239C28C7BD5D1CD76BFD282243430686BA6088A00C8BCE1E8FB2878C075448A`<br>**ACCDB MD5:** `6237949FA8E9F503C9623F0A7294310D`<br>**ACCDB SHA256:** `7A4B622211A66E7FE386291F9E1013CAF47088BC116498310A077338BE435196` | Implementación de `Property Get DB_PATH` con resolución en cascada (Local > Preprod > Dev) y caché `m_sRutaCacheBD` en `modDatabase.bas`. Redirección de llamadas en `AsignarNumFactura`, `ModuloLOG` y `ModuloSeguridadLocal`. |
| **3. Backup Preventivo PRE-SANEAMIENTO WAL**<br>*(Antes de Purga y Ajuste Idempotente)* | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2240_PRE_SANEAMIENTO_WAL.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2240_PRE_SANEAMIENTO_WAL.accdb` | 6.385.789 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `549921707952BC6EAA423A6CBC107500`<br>**XLSM SHA256:** `9239C28C7BD5D1CD76BFD282243430686BA6088A00C8BCE1E8FB2878C075448A`<br>**ACCDB MD5:** `6DD968C7E5E84213266B255FB950FA44`<br>**ACCDB SHA256:** `F4E4D8FED4950590207C6772BB6DBAF8AF51DED8706CCFF607143CAB19F41B9F` | Respaldo del libro y backend antes de abordar la no-idempotencia en `EliminarOrdenBD` y la depuración del búfer WAL saturado. |
| **4. Espejo Físico WAL Pre-Saneamiento** | `H:\ResidenciaBD\BackupsDiarios\SYS_COLA_SYNC_PRE_SANEAMIENTO.wal.bak` | 220.155 B | — | Copia de seguridad del archivo WAL físico local en `%LOCALAPPDATA%\ResidenciaApp\SICO_N2\SYS_COLA_SYNC_SALÓN-PC.wal` previa a su supresión. |
| **5. Backup POST-SANEAMIENTO WAL E IDEMPOTENCIA** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2250_POST_SANEAMIENTO_WAL.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2250_POST_SANEAMIENTO_WAL.accdb` | 6.277.209 B (XLSM)<br>6.746.112 B (ACCDB) | **XLSM MD5:** `BF866C8349F030DDF0246C8AC514C737`<br>**XLSM SHA256:** `7AA6DAF21AD69CE410567A2901F6B905175D94092F680F4F191262D1BC87E551`<br>**ACCDB MD5:** `6DD968C7E5E84213266B255FB950FA44`<br>**ACCDB SHA256:** `F4E4D8FED4950590207C6772BB6DBAF8AF51DED8706CCFF607143CAB19F41B9F` | `EliminarOrdenBD` devuelve `True` y `CommitTrans` cuando el registro no existe en Access (`rsCheck.EOF = True`). Purga de 1.749 líneas de `SYS_COLA_SYNC`. Reducción de >108 KB en el `.xlsm`. Barra de estado en frío: `[CONECTADO]`. |
| **6. Backup PRE-FIX PROTECCIÓN FECHAS** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2310_PRE_FIX_PROTECCION_FECHAS.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2310_PRE_FIX_PROTECCION_FECHAS.accdb` | 6.277.209 B (XLSM)<br>6.750.208 B (ACCDB) | **XLSM MD5:** `BF866C8349F030DDF0246C8AC514C737`<br>**XLSM SHA256:** `7AA6DAF21AD69CE410567A2901F6B905175D94092F680F4F191262D1BC87E551`<br>**ACCDB MD5:** `877BB89415ACE72DC8F9FD417B67FAB5`<br>**ACCDB SHA256:** `DFA15AFDF7ABB18F9BB204B4F0E90A0E8A76154C66AA7E0799AEC919A5920759` | Respaldo previo a subsanar el bloqueo emergente de hoja protegida al editar la Columna B (`FECHA PETICION`) en `RESIDENCIA OVIEDO`. |
| **7. Backup POST-FIX PROTECCIÓN FECHAS** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2315_POST_FIX_PROTECCION_FECHAS.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2315_POST_FIX_PROTECCION_FECHAS.accdb` | 6.286.240 B (XLSM)<br>6.750.208 B (ACCDB) | **XLSM MD5:** `F3BC41E860FD5EE4D03BC76F037843D4`<br>**XLSM SHA256:** `D55B4D34565128E70B2481A0A07D0341A6915C74450BA7FA46042AFD9F8AB3BE`<br>**ACCDB MD5:** `877BB89415ACE72DC8F9FD417B67FAB5`<br>**ACCDB SHA256:** `DFA15AFDF7ABB18F9BB204B4F0E90A0E8A76154C66AA7E0799AEC919A5920759` | Sustitución de contraseña vacía por `ModuloConfigSegura.ObtenerPasswordHojas()` en `FechasPeticion.bas`. Reprotección blindada en `CleanUp:` y `ErrHandler:` preservando `UserInterfaceOnly:=True` y permisos completos de formato y edición. |
| **8. Backup POST-FIX PROTECCIÓN LISTAS CP/POBLACIÓN** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-02_2340_POST_FIX_PROTECCION_LISTAS.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-02_2340_POST_FIX_PROTECCION_LISTAS.accdb` | 6.616.947 B (XLSM)<br>6.754.304 B (ACCDB) | **XLSM MD5:** `06E31C80A08574CA44F016850FBA6655`<br>**XLSM SHA256:** `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`<br>**ACCDB MD5:** `1A4836EAEBB42D5FAA3D6C06F8CDEBBA`<br>**ACCDB SHA256:** `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | Integración quirúrgica de `DesprotegerHojaSilenciosa` y `ReprotegerHojaSilenciosa` en `ListasDinamicasCPyPoblaciones.bas`. Erradicación del error de hoja protegida al interactuar con el desplegable de poblaciones (Col Z) y códigos postales en `RESIDENCIA OVIEDO`, `GIJÓN` y `SOTO`. |
| **9. Versión de Referencia Consolidada Final**<br>*(Hito Oficial 2026-11-03 00:15)* | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-03_0015.xlsm`<br>`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`<br>`H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.xlsm`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`<br>`H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-03_0015.accdb`<br>`H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-11-03_0015_VERSION_REFERENCIA.accdb` | 6.616.947 B (XLSM)<br>6.754.304 B (ACCDB) | **XLSM MD5:** `06E31C80A08574CA44F016850FBA6655`<br>**XLSM SHA256:** `B81B9AC061512A29137F3922D4F921E56DF39F4308B12588539440CD9FDBB8B9`<br>**ACCDB MD5:** `1A4836EAEBB42D5FAA3D6C06F8CDEBBA`<br>**ACCDB SHA256:** `987823B76CB6DADCA29BADB9EE93500DAD64E9C628554134DF893153A1290A15` | **Versión de Referencia Vigente Oficial.** Consolidación definitiva de la jornada del 02/11/2026. 0 errores de compilación VBE, 0 mojibake, conectividad multi-entorno y sincronización 100% estabilizada. |

---

## 2. Descripción Técnica Exhaustiva de las Correcciones Implementadas

### 2.1. Resolución Dinámica de Backend y Memoria Caché (`modDatabase.bas`)
- **Problema:** Módulos centrales (`modDatabase`, `AsignarNumFactura`, `ModuloLOG`, `ModuloSeguridadLocal`) utilizaban una constante estática de red (`H:\GIJ1\AYUDANTIA\01 RESIDENCIAS\BDAS Asturias\Residencia_BE.accdb`). Al ejecutarse en equipos de desarrollo local donde el recurso de red no existe, la sonda de red devolvía `False` y forzaba el sistema a caer en Modo Contingencia N2. Esto bloqueaba la consulta OLEDB de autocompletado de códigos postales y municipios.
- **Solución:** Se sustituyó la constante fija por la propiedad dinámica `Property Get DB_PATH` con resolución en cascada defensiva:
  1. Carpeta local del libro activo (`ThisWorkbook.Path & "\Residencia_BE.accdb"`).
  2. Red de terminales preproducción (`H:\GIJ1\...`).
  3. Carpeta de desarrollo local (`H:\ResidenciaBD\Residencia_BE.accdb`).
  Se incorporó una variable estática en memoria (`m_sRutaCacheBD`) para evitar consultas redundantes a disco vía `Scripting.FileSystemObject`, complementada por `LimpiarCacheRutaBD`.
- **Resultado:** Autocompletado inmediato de 14.602 códigos postales en cualquier entorno sin intervención de código ni reconfiguraciones manuales.

### 2.2. Idempotencia en Eliminación y Purga del Búfer WAL SICO-N2
- **Problema:** La aplicación iniciaba con el aviso persistente `[MODO CONTINGENCIA N2 - 1749 cambios en búfer local]`. La investigación técnica determinó que:
  1. Existían 1.749 operaciones de borrado acumuladas en la hoja oculta `SYS_COLA_SYNC` y en el archivo `.wal` local.
  2. Al procesar las bajas hacia Access, `modDatabase.EliminarOrdenBD` verificaba la existencia previa del registro:
     ```vba
     If Not rsCheck.EOF Then
         ' Bloqueado
     Else
         rsCheck.Close
         cn.RollbackTrans
         EliminarOrdenBD = False  ' <--- ERROR DE IDEMPOTENCIA
         Exit Function
     End If
     ```
  3. Como las órdenes ya habían sido borradas físicamente en pruebas anteriores, la función devolvía `False`. El encolador SICO-N2 interpretaba el retorno como un fallo de red transitorio, no purgaba la cola y marcaba el sistema en contingencia offline permanente.
- **Solución:** Se aplicó el principio matemático de idempotencia ($f(f(x)) = f(x)$): en colas de sincronización distribuida, solicitar la eliminación de un registro que ya no existe constituye un éxito funcional:
  ```vba
  Else
      rsCheck.Close
      Set rsCheck = Nothing
      If enTransaccion Then cn.CommitTrans
      enTransaccion = False
      cn.Close
      Set cn = Nothing
      EliminarOrdenBD = True ' Idempotente: el estado final deseado ya se cumple
      Exit Function
  End If
  ```
  Se purgaron físicamente las 1.749 líneas residuales de `SYS_COLA_SYNC` y se eliminó el archivo residual desfasado de 220 KB en `%LOCALAPPDATA%\...` (previamente respaldado).
- **Resultado:** `ContarCambiosPendientes() = 0`, el libro `.xlsm` se aligeró en más de **108 KB** y la aplicación conecta en frío instantáneamente en `[CONECTADO]`.

### 2.3. Blindaje de Protección en Columna B (`FechasPeticion.bas`)
- **Problema:** Al introducir o editar la fecha de petición en la columna B de `RESIDENCIA OVIEDO`, saltaba un cuadro de diálogo: *"No se puede modificar las celdas porque la hoja está protegida"*. El valor permanecía en la cuadrícula pero la macro abortaba, omitiendo el formateo canónico (`dd/mm/yyyy hh:mm`) y la persistencia hacia Access.
- **Solución:** Se detectó que `FechasPeticion.bas` ejecutaba `hoja.Unprotect password:=""` (contraseña vacía hardcodeada), la cual fallaba contra la contraseña maestra del proyecto. Además, en los bloques de salida (`CleanUp:` y `ErrHandler:`) se ejecutaba `hoja.Protect password:=""` sin argumentos, lo que despojaba a la hoja de la directiva volátil `UserInterfaceOnly:=True` y de los permisos de usuario.
  Se refactorizó el módulo enlazándolo con `ModuloConfigSegura.ObtenerPasswordHojas()` y reprogramando la reprotección estricta garantizando `UserInterfaceOnly:=True` y permisos completos de formato y edición.

### 2.4. Resolución de Conflicto en Desplegables de Población (`ListasDinamicasCPyPoblaciones.bas`)
- **Problema:** Al interactuar con la columna Z (POBLACIÓN) en `RESIDENCIA OVIEDO`, saltaba de forma idéntica el emergente de hoja protegida.
- **Solución:** En Excel, los métodos `.Validation.Delete` y `.Validation.Add` arrojan un error en tiempo de ejecución si la hoja está protegida, aun cuando la celda tenga `Locked = False`. Las subrutinas `CrearValidacionListaSilenciosa`, `CrearValidacionAutocompletadoOptimizada` y `LimpiarValidacionSilenciosa` llamaban a `celda.Worksheet.Unprotect` sin contraseña y `celda.Worksheet.Protect` sin parámetros.
  Se incorporaron las subrutinas privadas auxiliares `DesprotegerHojaSilenciosa` y `ReprotegerHojaSilenciosa`:
  - Consumen la contraseña maestra `ModuloConfigSegura.ObtenerPasswordHojas()`.
  - Aseguran la reaplicación de `UserInterfaceOnly:=True` y de todos los flags de permisos operativos (`AllowFormattingCells`, `AllowSorting`, etc.).

### 2.5. Validación Desacoplada de Capacidad Habitacional (`Hoja2.cls` vs `Hoja24.cls`)
- **Problema:** Incoherencia en la validación entre PAX y habitaciones individuales/dobles en `RESIDENCIA GIJÓN`.
- **Solución:** Se adecuó la lógica a la estructura específica de cada sede:
  - En `RESIDENCIA GIJÓN`: Columna S representa el número de habitaciones adjudicadas, verificándose directamente $Q + R = S$.
  - En `RESIDENCIA OVIEDO`: Columna S corresponde a camas supletorias (condicionadas a $R > 0$ y con límite físico) y Columna T a las habitaciones adjudicadas.

---

## 3. Estado de Calidad y Verificación

- **Compilación VBE:** 0 errores (verificado mediante automatización COM Excel).
- **Control de Mojibake:** 0 caracteres corruptos detectados en el escaneo automatizado contra patrones `Ã`, `Â`, `ǭ`, `?`. Codificación 100% nativa **ANSI Windows-1252**.
- **Regla Estricta de Sede:** 100% conforme con el texto literal **`GIJÓN`** (prohibido `Chr(211)`).
- **Pruebas de Usuario:** Confirmadas y validadas por el operador (*"Compila sin errores y funciona según lo esperado"*).

