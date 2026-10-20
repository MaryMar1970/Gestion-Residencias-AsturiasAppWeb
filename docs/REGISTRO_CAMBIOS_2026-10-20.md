# Registro Diario de Control de Cambios — 20/10/2026

> **Entorno de Trabajo Activo:** `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
> **Servidor Compartido en Red:** `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
> **Base de Datos Backend:** `H:\ResidenciaBD\Residencia_BE.accdb`  
> **Versión de Referencia Oficial:** `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (Hito 2026-10-20 19:35)  
> **Estado de Compilación VBA:** 🟢 **0 ERRORES** (Validado en Microsoft Excel 2016 COM Automation)  
> **Estado de Codificación (Mojibake):** 🟢 **0 ARCHIVOS / 0 COMPONENTES AFECTADOS** (100% Windows-1252 ANSI Limpio)  

---

## 1. Resumen Ejecutivo de la Sesión

La sesión de trabajo del 20 de octubre de 2026 culminó con la resolución integral de las anomalías en el sistema de auditoría LOG en las tres residencias, la eliminación de bloqueos en recálculos de hoja y el blindaje permanente de la sesión y el rol de usuario:

1. **Saneamiento Estructural de Hojas LOG (`LOG_GIJÓN`, `LOG_SOTO`, `LOG_OVIEDO`)**:
   - Corrección de la separación y parseo de columnas en [ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas).
   - Volcado exacto en columnas A a F (*Usuario*, *Fecha/Hora*, *Nº Orden*, *Columna/Acción*, *Valor anterior*, *Valor nuevo*).
2. **Sincronización Relacional Invertida de los 500 Registros Más Recientes**:
   - Corrección de la consulta SQL hacia la tabla `LogActividad` de Access para traer los últimos 500 eventos en lugar de los más antiguos.
3. **Desambiguación de Eventos `LOGIN` por Residencia**:
   - Filtrado estricto por residencia para evitar que logins de Gijón o Soto aparezcan en la hoja de Oviedo.
4. **Blindaje contra Error 1004 en Fórmulas Auditadas**:
   - Tratamiento de cadenas que comienzan por `=` o `+` insertándolas con apóstrofo como texto puro.
5. **Erradicación del Error 50290 en `Worksheet_Calculate` de `Hoja11` (`SOBRE GIJÓN`)**:
   - Eliminación de `Application.EnableEvents = True/False` durante el evento `Calculate` al interactuar con desplegables (como la columna F de Situación en Gijón).
6. **Persistencia Auto-reparable de Sesión y Rol (ADMIN/OPERADOR)**:
   - Mecanismo en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) mediante Nombres Ocultos de Excel para rehidratar credenciales si se resetea la memoria de VBA.
7. **Blindaje Permanente de la Barra de Estado**:
   - Preservación de la leyenda informativa en la barra inferior sin que sea reemplazada por avisos de accesibilidad al cambiar de pestaña o usar el Ribbon.
8. **Establecimiento de la Nueva Versión de Referencia Oficial**:
   - Respaldo incremental, archivado del hito previo a `Historico/`, protección de sólo lectura y sincronización con el servidor de red.

---

## 2. Detalle de Incidencias y Requerimientos Abordados

### 2.1. Saneamiento del Parseo y Volcado en Hojas LOG
- **Problema detectado:** Al grabarse los cambios en las hojas LOG, los campos no se distribuían en las columnas adecuadas:
  - Columna C (*Nº ORDEN*) quedaba vacía.
  - Columna D (*COLUMNA*) recibía `LOGIN` o `CAMBIO (GIJON)`.
  - Columna E (*Valor anterior*) quedaba vacía.
  - Columna F (*Valor nuevo*) recibía una cadena concatenada: `Nº Orden: 640 | Campo: Nº ORDEN | Antes:  | Ahora: 640`.
- **Causa raíz:** En [ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas), el procedimiento `SincronizarHojaLOGDesdeAccess` asumía un formato de detalle monolítico y no parseaba las etiquetas internas `" | Campo: "`, `" | Antes: "` y `" | Ahora: "`.
- **Solución implementada:**
  - Rediseño del parseador de cadenas en `SincronizarHojaLOGDesdeAccess` para extraer atómicamente cada variable y escribirla en su respectiva columna:
    - Columna A: Usuario.
    - Columna B: Fecha/Hora formateada (`dd/mm/yyyy hh:mm`).
    - Columna C: Nº de Orden.
    - Columna D: Columna modificada o acción.
    - Columna E: Valor anterior limpio.
    - Columna F: Valor nuevo limpio.

### 2.2. Sincronización de los 500 Registros Más Recientes (SQL Invertido)
- **Problema detectado:** Las hojas LOG de Gijón y Soto solo mostraban datos anclados en fechas antiguas (22/09/2026 y 02/10/2026) a pesar de introducir nuevas reservas.
- **Causa raíz:** La consulta ejecutaba `SELECT TOP 500 * FROM LogActividad ORDER BY Id ASC`. Con más de 8.350 registros en Access, siempre traía los 500 primeros registros creados en septiembre, truncando toda la actividad posterior.
- **Solución implementada:** Sustitución por una subconsulta SQL anidada:
  ```sql
  SELECT * FROM (
      SELECT TOP 500 * FROM LogActividad 
      WHERE Residencia = 'GIJON' OR (Accion = 'LOGIN' AND Detalle LIKE '%GIJON%')
      ORDER BY Id DESC
  ) ORDER BY Id ASC
  ```
  De este modo, se seleccionan los 500 registros más recientes y se ordenan cronológicamente para su volcado en Excel.

### 2.3. Desambiguación de Logins y Causa de Bloqueo en Oviedo
- **Problema detectado:** La hoja `LOG_OVIEDO` no avanzaba más allá del 05/10/2026 21:52 y mostraba eventos de inicio de sesión de Gijón.
- **Causa raíz doble:**
  1. La cláusula `OR Accion = 'LOGIN'` sin filtro de residencia contaminaba el log de Oviedo con accesos de otros centros.
  2. En el registro #247 existía una fórmula compleja (`=SI(O...`). Al intentar volcarla directamente con `.Value = nuevVal`, Excel intentaba evaluarla como fórmula y lanzaba un Error en tiempo de ejecución 1004, abortando la sincronización en seco.
- **Solución implementada:**
  - Filtrado específico: `OR (Accion = 'LOGIN' AND Detalle LIKE '%OVIEDO%')`.
  - Tratamiento seguro de valores de texto: si `Left$(nuevVal, 1) = "="` o `Left$(antVal, 1) = "="`, se antepone `'` para forzar su interpretación como texto literal sin ejecutar evaluación de fórmula.

### 2.4. Erradicación del Error 50290 en `Hoja11` (`SOBRE GIJÓN`)
- **Problema detectado:** Al seleccionar una opción en la lista desplegable de la columna F (*SITUACIÓN*) en `RESIDENCIA GIJÓN`, la hoja se bloqueaba con el error emergente 50290 en tiempo de ejecución: *Error en el método EnableEvents de objeto _Application*, resaltando en amarillo `Application.EnableEvents = True` en `Private Sub Worksheet_Calculate` de `Hoja11`.
- **Causa raíz:** Durante la fase de recálculo (`Worksheet_Calculate`), el motor interno de Excel prohíbe taxativamente alterar la propiedad `EnableEvents`. Cualquier intento de modificarla arroja de inmediato el error 50290, dejando la interfaz gráfica congelada a medio redibujar.
- **Solución implementada:** Supresión completa de las manipulaciones de `EnableEvents` dentro de `Hoja11.Worksheet_Calculate`, envolviendo la rutina en un bloque seguro con `On Error Resume Next`.

### 2.5. Persistencia de Sesión, Rol ADMIN y Blindaje de Barra de Estado
- **Problema detectado:** Al iniciar sesión con rol Administrador, la barra de estado mostraba correctamente la conexión. Sin embargo, al pulsar un botón del Ribbon o cambiar de pestaña, la leyenda desaparecía y se mostraba el aviso de Office: *"Accesibilidad: es necesario investigar"*. A partir de ese momento, las hojas LOG no registraban el usuario y se perdía el rol ADMIN, impidiendo el desbloqueo de hojas.
- **Causa raíz:**
  1. En los callbacks del Ribbon se ejecutaba `Application.StatusBar = False`, devolviendo el control de la barra a Excel y provocando que Office 365 mostrara sus alertas de accesibilidad.
  2. Cualquier error fortuito o recompilación de macros reseteaba las variables volátiles en memoria VBA (`m_NombreUsuario = ""`, `m_Rol = ""`, `ThisWorkbook.usuarioActual = ""`), haciendo que `EsAdmin()` retornara `False`.
- **Solución implementada:**
  - **Almacén persistente en Nombres Definidos ocultos:** Se implementaron `GuardarSesionEnLibro` y `RehidratarSesionSiEsNecesario` en [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas), guardando el usuario y rol en `ThisWorkbook.Names("_Session_Usuario")`, etc.
  - **Rehidratación automática:** En `Workbook_SheetActivate` y antes de acciones sensibles, si las variables en memoria están vacías, se leen de inmediato de los nombres definidos.
  - **Preservación de Barra:** Sustitución de `StatusBar = False` por `modDatabase.ActualizarBarraEstado()`, garantizando que la leyenda del usuario conectado se mantenga siempre visible y activa.

---


### 2.6. Tarifas Dinámicas en Facturas PAX de Gijón y Oviedo vinculadas a TARIFAS UNIDAD
- **Problema detectado:** Al desdoblar en la hoja `TARIFAS UNIDAD` las tarifas vigentes por periodos temporales (Gijón en `$E$3:$O$9` y Oviedo en `$E$13:$O$19`), las fórmulas de cálculo de precio diario en las facturas físicas imprimibles apuntaban a celdas fijas de la fila 3 (Gijón) y fila 13 (Oviedo) sin ponderar la fecha de estancia:
  - En `FACTURA GIJÓN`: `R23` (Hab. Individual) apuntaba a `'TARIFAS UNIDAD'!G3` y `'TARIFAS UNIDAD'!H3`; `R24` (Hab. Doble) apuntaba a `I3` y `J3`.
  - En `FACTURA OVIEDO`: `R23` apuntaba a `G13` y `H13`; `R24` a `I13` y `J13`; `R25` (Cama Supletoria) a `K13` y `L13`.
  - Como consecuencia, cualquier cambio de tarifa anual (ej. tarifas 2026 respecto a 2025) provocaba que la hoja de residencia calculara el nuevo importe correcto en columna U/V, pero la factura impresa siguiese liquidando con el precio fijo del ejercicio anterior.
- **Causa raíz:** Las plantillas de factura fueron concebidas originalmente con celdas estáticas simples previas a la incorporación de la matriz dinámica multi-ejercicio gobernada por anti-solapes en `TARIFAS UNIDAD`.
- **Solución implementada:**
  - Sustitución de las referencias estáticas por fórmulas dinámicas matriciales mediante `SUMAPRODUCTO`, evaluando la Fecha de Entrada (`F19`, equivalente a `L2` de las hojas de residencia) y aplicando el filtro de celda no vacía `('TARIFAS UNIDAD'!$E$...>0)`:
    - **`FACTURA GIJÓN` (`R23` - Hab. Individual):**
      `=((SI(I23="HABITACION INDIVIDUAL";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=F19)*('TARIFAS UNIDAD'!$F$3:$F$9>=F19)*'TARIFAS UNIDAD'!$G$3:$G$9);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA GIJÓN'!$A:$AE;8;0);"")="NO";1;0))+(SI(I23="HABITACION INDIVIDUAL";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=F19)*('TARIFAS UNIDAD'!$F$3:$F$9>=F19)*'TARIFAS UNIDAD'!$H$3:$H$9);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA GIJÓN'!$A:$AE;8;0);"")="SI";1;0)))*Z14`
    - **`FACTURA GIJÓN` (`R24` - Hab. Doble):**
      `=((SI(I24="HABITACION DOBLE";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=F19)*('TARIFAS UNIDAD'!$F$3:$F$9>=F19)*'TARIFAS UNIDAD'!$I$3:$I$9);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA GIJÓN'!$A:$AE;8;0);"")="NO";1;0))+(SI(I24="HABITACION DOBLE";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=F19)*('TARIFAS UNIDAD'!$F$3:$F$9>=F19)*'TARIFAS UNIDAD'!$J$3:$J$9);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA GIJÓN'!$A:$AE;8;0);"")="SI";1;0)))*Z14`
    - **`FACTURA OVIEDO` (`R23` - Hab. Individual):**
      `=((SI(I23="HABITACION INDIVIDUAL";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$G$13:$G$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="NO";1;0))+(SI(I23="HABITACION INDIVIDUAL";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$H$13:$H$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="SI";1;0)))*$AB$1`
    - **`FACTURA OVIEDO` (`R24` - Hab. Doble):**
      `=((SI(I24="HABITACION DOBLE";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$I$13:$I$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="NO";1;0))+(SI(I24="HABITACION DOBLE";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$J$13:$J$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="SI";1;0)))*$AB$1`
    - **`FACTURA OVIEDO` (`R25` - Cama Supletoria):**
      `=((SI(I25="CAMA SUPLETORIA";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$K$13:$K$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="NO";1;0))+(SI(I25="CAMA SUPLETORIA";SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$13:$E$19>0)*('TARIFAS UNIDAD'!$E$13:$E$19<=F19)*('TARIFAS UNIDAD'!$F$13:$F$19>=F19)*'TARIFAS UNIDAD'!$L$13:$L$19);0))*(SI(SI.ERROR(BUSCARV(D1;'RESIDENCIA OVIEDO'!$A:$AE;8;FALSO);"")="SI";1;0)))*$AB$1`
  - Se verificó mediante auditoría forense que ningún otro elemento del libro (las 49 hojas, 90 nombres de rango y 144 módulos VBA) requiere ajustes, quedando el circuito contable perfectamente coordinado y cerrado.

## 3. Trazabilidad Criptográfica de Archivos y Respaldos

### 3.1. Versión de Referencia Oficial (Hito 2026-10-20 19:35)
- **Directorio Protegido:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`
  - `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
    - **MD5:** `135fdb2285f4983e232d59c76c0074e2`
    - **Tamaño:** `7.152.489 bytes`
    - **Atributo:** `IsReadOnly = True`
  - `Residencia_BE_VERSION_REFERENCIA.accdb`
    - **MD5:** `c40aa4331f3b83acacd5132707ed6baa`
    - **Tamaño:** `6.721.536 bytes`
    - **Atributo:** `IsReadOnly = True`

### 3.2. Copia Incremental Asociada
- **Directorio de Backups:** `H:\ResidenciaBD\BackupsDiarios\`
  - `BDAS_Multiusuario_BACKUP_2026-10-20_1935_TARIFAS_DINAMICAS_FACTURAS.xlsm` (MD5: `135fdb2285f4983e232d59c76c0074e2`)
  - `Residencia_BE_BACKUP_2026-10-20_1935_TARIFAS_DINAMICAS_FACTURAS.accdb` (MD5: `c40aa4331f3b83acacd5132707ed6baa`)

### 3.3. Archivos Históricos Archivados
- **Directorio Histórico:** `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_0100.xlsm` (MD5: `1b7b1950f5071d2927ed31bebc7b2f5d`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-10-20_0100.accdb` (MD5: `54522630b0bb8d2229eda29029e22bfd`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_2135.xlsm` (MD5: `57821d0835bcbb82b5356a722b982945`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-10-19_2135.accdb` (MD5: `dd7c3f0dc0d62b4daef487de9d2f303d`)

### 3.4. Archivos Activos y Despliegue de Red Sincronizados
- `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` ➔ `135fdb2285f4983e232d59c76c0074e2`
- `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` ➔ `135fdb2285f4983e232d59c76c0074e2`
- `H:\ResidenciaBD\Residencia_BE.accdb` ➔ `c40aa4331f3b83acacd5132707ed6baa`

## 4. Estado de Validación Técnica

- **Compilación VBA en Excel 2016 COM:** 🟢 **0 ERRORES** (`Compile VBAProject` limpio y deshabilitado).
- **Linter de Codificación y Mojibake:** 🟢 **127 archivos analizados / 0 afectados / 0 problemas** (100% Windows-1252 ANSI).
- **Regla Estricta de Nombres:** Cumplimiento total del texto literal `GIJÓN` (`0` ocurrencias de `Chr(211)`).
- **Golden Master:** Intacto e inalterado en `H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm`.
