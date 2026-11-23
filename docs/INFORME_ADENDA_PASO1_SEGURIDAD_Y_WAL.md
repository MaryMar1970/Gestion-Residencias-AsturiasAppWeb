# RESIDENCIAAPP — PUERTA 2: LABORATORIO TÉCNICO
## ADENDA DE SEGURIDAD, CONEXIÓN REAL Y AISLAMIENTO OBLIGATORIO DE SICO-N2/WAL

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Puerta 2 — Laboratorio Técnico v1.0  
**Hito:** Cierre de comprobaciones complementarias de seguridad y precondición WAL  
**Documento base:** [`INFORME_PASO1_SETUP_LABORATORIO.md`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/EVIDENCIAS/INFORME_PASO1_SETUP_LABORATORIO.md)  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha y hora:** 20 de noviembre de 2026 | 16:12  

---

## 1. PERMISOS NTFS EFECTIVOS Y PROTECCIÓN DE DATOS OPERATIVOS

En cumplimiento de las directrices de seguridad, se revisaron y endurecieron los permisos NTFS mediante `icacls` exclusivamente dentro del directorio del laboratorio, sin alterar el origen ni ninguna carpeta de producción:

### 1.1. Configuración de Permisos por Directorio

| Directorio | Permisos NTFS Efectivos (`icacls`) | Justificación Técnica y Nivel de Aislamiento |
| :--- | :--- | :--- |
| **`LABORATORIO_GATE2_...` (Raíz)** | `Administradores:(F)`<br>`SYSTEM:(F)`<br>`Usuarios autentificados:(M)`<br>`Usuarios:(RX)` | Entorno local en `SALÓN-PC`. No existen recursos compartidos ni exposición de red a terceros. |
| **`BASELINE\`** | `Administradores:(OI)(CI)(F)`<br>`SYSTEM:(OI)(CI)(F)`<br>`salón-pc\jaime_ntymfau:(OI)(CI)(RX)` | **Herencia deshabilitada (`/inheritance:r`).** Se eliminaron por completo `Usuarios autentificados` y `Usuarios`. La cuenta del ejecutor tiene **exclusivamente lectura y ejecución (`RX`)**, sin permisos de modificación, escritura ni eliminación. |
| **`BASELINE\*` (Ficheros)** | `Administradores:(F)`<br>`SYSTEM:(F)`<br>`salón-pc\jaime_ntymfau:(RX)` | Además de la ACL restrictiva, ambos ficheros conservan el atributo del sistema de archivos `+R` (Solo Lectura). |
| **`TRABAJO\`** | `Administradores:(F)`<br>`SYSTEM:(F)`<br>`salón-pc\jaime_ntymfau:(M)` | Permite lectura, creación y modificación de esquemas únicamente a la cuenta autorizada para las pruebas. |
| **`EVIDENCIAS\`** | `Administradores:(F)`<br>`SYSTEM:(F)`<br>`salón-pc\jaime_ntymfau:(M)` | Permite el volcado de logs, manifiestos JSON y capturas de auditoría. |

### 1.2. Prueba Empírica de Bloqueo en BASELINE

Se ejecutó un intento automatizado de escritura en [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\BASELINE\test_write.txt`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/BASELINE/test_write.txt):
```text
[OK] PermissionError verificado en BASELINE (escritura bloqueada por ACL/R):
     [Errno 13] Permission denied: 'H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\BASELINE\test_write.txt'
```
**Conclusión:** La inmutabilidad de `BASELINE` queda blindada tanto por atributos de fichero como por listas de control de acceso (ACLs) del sistema de archivos NTFS.

---

## 2. VERIFICACIÓN DE CONEXIÓN REAL DE SOLO LECTURA (DAO Y ADODB 32 BITS)

Se ejecutó una prueba de conexión efectiva contra [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\Residencia_BE.accdb`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/TRABAJO/Residencia_BE.accdb) utilizando un proceso de 32 bits (`SysWOW64\cscript.exe`) para verificar la apertura, la ruta física real, la lectura y la ausencia de bloqueos.

### 2.1. Ensayo 1: Motor DAO 32 bits (`DAO.DBEngine.120`)

1. **Apertura:** `dao.OpenDatabase(labDbPath, False, True)` (modo `ReadOnly = True`).
2. **Ruta física informada por el motor (`db.Name`):**  
   `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\Residencia_BE.accdb`
3. **Estado de bloqueos durante la conexión:**
   - `TRABAJO\Residencia_BE.laccdb`: **`Verdadero`** *(demuestra que la conexión abrió físicamente el archivo local)*.
   - `H:\ResidenciaBD\Residencia_BE.laccdb`: **`Falso`** *(cero conexiones al origen de producción)*.
4. **Consulta no destructiva:** `SELECT COUNT(*) AS Total FROM Ordenes` $\rightarrow$ **2.127 registros** leídos con éxito.
5. **Cierre:** `db.Close` y liberación del objeto COM.

### 2.2. Ensayo 2: Motor ADODB OLEDB 32 bits (`Microsoft.ACE.OLEDB.12.0`)

1. **Apertura:** `conn.Mode = 1` (`adModeRead`), cadena de conexión apuntando a `TRABAJO`.
2. **Estado de bloqueos durante la conexión:**
   - `TRABAJO\Residencia_BE.laccdb`: **`Verdadero`**.
   - `H:\ResidenciaBD\Residencia_BE.laccdb`: **`Falso`**.
3. **Consulta no destructiva:** `SELECT COUNT(*) AS Total FROM ConfigTarifasCeldas` $\rightarrow$ **291 registros** leídos con éxito.
4. **Cierre:** `conn.Close` y liberación del objeto.

### 2.3. Comprobación Posterior de Bloqueos Residuales

Inmediatamente tras el cierre de ambas conexiones:
- `Existe laccdb en TRABAJO (post):` **`Falso`** (liberado y eliminado automáticamente por el motor).
- `Existe laccdb en PRODUCCION (post):` **`Falso`** (nunca fue tocado).
- **Resultado:** **PASS TOTAL — Cero bloqueos residuales y aislamiento físico comprobado**.

---

## 3. EXPLICACIÓN TÉCNICA DEL DIAGNÓSTICO ANTERIOR (`test_db_path.vbs`)

- **Causa del código de salida 1 en `test_db_path.vbs`:**  
  En [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L63), `DB_PATH` está declarado como una propiedad:  
  `Public Property Get DB_PATH() As String`.  
  En el modelo de objetos de automatización COM de Excel, el método `Application.Run` está diseñado por especificación para ejecutar únicamente procedimientos de tipo `Sub` o `Function`. Al intentar ejecutar `xl.Run("DB_PATH")` o `xl.Run("modDatabase.DB_PATH")`, Excel lanza el error de tiempo de ejecución:  
  *«No se puede ejecutar la macro "DB_PATH". Puede que la macro no esté disponible en este libro o que se hayan deshabilitado todas las macros.»*  
  El script `test_db_path.vbs` capturó esta excepción esperada y finalizó con código 1.

- **Utilización posterior de `test_get_connection_string.vbs`:**  
  Para resolver la invocación sin alterar el código del libro, se localizó en [`modDatabase.bas#L682`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L682) la función:
  ```vba
  Public Function GetConnectionString() As String
      GetConnectionString = PROVIDER_STRING & DB_PATH & ";"
  End Function
  ```
  Al ser una `Function` pública ordinaria, `Application.Run("GetConnectionString")` es plenamente ejecutable vía COM. Su llamada evaluó internamente la propiedad `DB_PATH` y devolvió la cadena completa con salida exitosa (código 0), acreditando que la resolución conducía exactamente a `TRABAJO\Residencia_BE.accdb`.

---

## 4. RESOLUCIÓN TÉCNICA DEL AISLAMIENTO SICO-N2 / WAL

### 4.1. Auditoría del Código de Gestión WAL en `ModuloSeguridadLocal.bas`

Se inspeccionaron todos los puntos de construcción, lectura, escritura y eliminación del diario WAL:
- [`VolcarColaHaciaDiscoWAL`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1460): Obtiene la ruta física llamando a `ObtenerRutaArchivoWAL()` y escribe el archivo `.wal` secuencialmente con delimitadores `|`. Si la cola está vacía, ejecuta `Kill rutaWAL`.
- [`RehidratarColaDesdeDiscoWALSiNecesario`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1499): Obtiene la ruta llamando a `ObtenerRutaArchivoWAL()`, comprueba `Dir$(rutaWAL)` y rehidrata la hoja `SYS_COLA_SYNC`.
- [`ObtenerCarpetaLocalSICO`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1571): Punto centralizado de resolución de la carpeta del diario. En producción resuelve `%LOCALAPPDATA%\ResidenciaApp\SICO_N2`.
- [`ObtenerRutaArchivoWAL`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L1582): Concatena la carpeta anterior con `\SYS_COLA_SYNC_<COMPUTERNAME>.wal`.

### 4.2. Adaptación Quirúrgica en la Copia de Trabajo

En estricto cumplimiento de los requisitos 2, 3 y 4 del mandato, se modificó **exclusivamente la copia de trabajo** [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\BDAS_Multiusuario.xlsm`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/TRABAJO/BDAS_Multiusuario.xlsm) (preservando intacto `BASELINE`, el repositorio Git y producción) para redirigir la resolución de ruta:

```vba
Public Function ObtenerCarpetaLocalSICO() As String
    Dim targetDir As String, subDir As String
    ' Redirección exclusiva al directorio de laboratorio
    targetDir = ThisWorkbook.Path & "\WAL_TEST"
    If Len(Dir$(targetDir, vbDirectory)) = 0 Then On Error Resume Next: MkDir targetDir: On Error GoTo 0
    ' Soporte de aislamiento independiente por escenario
    subDir = Environ$("SICO_WAL_SUBDIR")
    If Len(subDir) > 0 Then
        targetDir = targetDir & "\" & subDir
        If Len(Dir$(targetDir, vbDirectory)) = 0 Then On Error Resume Next: MkDir targetDir: On Error GoTo 0
    End If
    ObtenerCarpetaLocalSICO = targetDir
End Function

Public Function ObtenerRutaArchivoWAL() As String
    ObtenerRutaArchivoWAL = ObtenerCarpetaLocalSICO() & "\SYS_COLA_SYNC_" & Environ$("COMPUTERNAME") & ".wal"
End Function
```

**Garantías Técnicas de la Modificación:**
1. **Aislamiento Total:** El diario se ubica en [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\WAL_TEST\`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/TRABAJO/WAL_TEST/).
2. **Soporte Multi-escenario:** Mediante la variable de entorno `SICO_WAL_SUBDIR`, cada prueba puede aislarse en su propia subcarpeta (ej. `WAL_TEST\ESC_01_RECUPERACION\`), permitiendo cerrar y reabrir Excel sobre el mismo WAL sin interferencias.
3. **Cero Alteración Lógica:** No se modificó ningún delimitador, cálculo CRC16, orden FIFO ni formato de registro.
4. **Comprobación en Runtime:** Se ejecutó `xl.Run("ObtenerRutaArchivoWAL")` desde Excel 2016 32-bit sobre el libro adaptado, obteniendo como resultado verificado:
   ```text
   H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\WAL_TEST\SYS_COLA_SYNC_SALÓN-PC.wal
   ```

---

## 5. EVIDENCIA DE INTEGRIDAD DE BASELINE, PRODUCCIÓN Y PERFIL WINDOWS

Se verificó el estado de todos los ficheros tras la ejecución de las comprobaciones complementarias:

| Fichero / Entorno | Ubicación | Hash SHA-256 Actual | Estado vs. Referencia |
| :--- | :--- | :--- | :---: |
| **Producción ACCDB** | `H:\ResidenciaBD\Residencia_BE.accdb` | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **100% INTACTO** |
| **Producción XLSM** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **100% INTACTO** |
| **BASELINE ACCDB** | `...LABORATORIO...\BASELINE\Residencia_BE.accdb` | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **100% INTACTO** |
| **BASELINE XLSM** | `...LABORATORIO...\BASELINE\BDAS_Multiusuario.xlsm` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **100% INTACTO** |
| **Perfil Real SICO-N2** | `C:\Users\jaime_ntymfau\AppData\Local\ResidenciaApp\SICO_N2\` | *(3 archivos históricos)* | **100% INTACTO** |
| - `wal.bak_pre_reintegrar` | `...SICO_N2\SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_reintegrar_...` | `20CD431C01719022F6A2C89B53F8521C48C31059FED90CBBC48FB56CC07DA119` | **100% INTACTO** |
| - `wal.bak_pre_revert` | `...SICO_N2\SYS_COLA_SYNC_SALÓN-PC.wal.bak_pre_revert_oviedo` | `892C19CE4F9C652151B77609DFFD6768D98057D70AE518FE189633FCAD1FA3E1` | **100% INTACTO** |
| - `cred_cache` | `...SICO_N2\cred_cache_SALÓN-PC.dat` | `4B94D58E1A327C07AE6AF467B19F52E2F1860D5DBEB5702316BE46A5D7076FF4` | **100% INTACTO** |

---

### Conclusión y Detención Formal

Todas las precondiciones de seguridad de la Puerta 2 han quedado formalmente acreditadas:
1. Permisos NTFS y protección de `BASELINE` operativos y verificados.
2. Conexión real 32 bits a `TRABAJO` demostrada sin impacto en producción y con cero bloqueos residuales.
3. Causa del diagnóstico VBS aclarada.
4. Redirección de SICO-N2/WAL a `TRABAJO\WAL_TEST\` establecida y comprobada en runtime sin tocar el perfil de Windows.

**El inspector técnico detiene su actuación en este punto.** No se iniciará ninguna prueba funcional, transaccional ni de escritura en `TRABAJO` hasta recibir la autorización expresa de la Dirección Arquitectónica.