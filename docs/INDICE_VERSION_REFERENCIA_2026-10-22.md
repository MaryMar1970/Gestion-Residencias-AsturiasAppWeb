# Índice Técnico de la Versión de Referencia Oficial — Hito 2026-10-22 19:45

> **Ubicación de Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`  
> **Libro Maestro:** `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (MD5: `DB8FFA1648DD2917C385E2AAED7E9BD8` | 6.755.115 bytes)  
> **Base de Datos Maestro:** `Residencia_BE_VERSION_REFERENCIA.accdb` (MD5: `6E41A8521D884B740907CD9DEF187CEA` | 6.729.728 bytes)  
> **Estado de Atributos:** Ambos archivos protegidos contra escritura (`IsReadOnly = True`).  
> **Script de Recuperación Inmediata:** `bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1`  

---

## 1. Alcance Funcional y Técnico Consolidado en este Hito

Esta versión de referencia consolida la totalidad de mejoras del hito previo (2026-10-21 23:30) e incorpora las siguientes soluciones y blindajes técnicos:

1. **Restauración Dinámica e Integral de las Barras de Desplazamiento (`ThisWorkbook.cls`)**:
   - **Ocultación exclusiva durante el Login:** En la pantalla de bienvenida (`INICIO`), las barras de desplazamiento vertical y horizontal se ocultan para mantener la carátula limpia y sin recorridos parásitos.
   - **Restauración completa post-arranque:** Inmediatamente tras completarse el arranque en `Workbook_Open`, se reponen de forma explícita las barras en la ventana activa (`ActiveWindow.DisplayVerticalScrollBar = True`, `ActiveWindow.DisplayHorizontalScrollBar = True`, `DisplayHeadings = True`, `DisplayGridlines = True`).
   - **Gobernanza reactiva por evento de hoja (`Workbook_SheetActivate`):** Se intercepta el evento de cambio de hoja de modo que cualquier hoja operativa (`RESIDENCIA OVIEDO`, `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO`, `CALENDARIO ...`, `RESUMEN ...`, `LOG ...`, etc.) mantenga siempre visibles las barras de scroll, permitiendo el desplazamiento fluido a través de todas las columnas (más allá de la columna S y AE) y filas.
   - **Restauración en contingencia (`CleanClose`):** Si el login se cancela o falla, las barras se restauran antes del cierre para evitar estados degradados de Excel.

2. **Saneamiento de Módulo de Navegación por Columnas (`BotonIzdaDchaHojasRESIDENCIAS.bas`)**:
   - Supresión de caracteres corruptos mojibake (`GIJï¿½N` $\to$ `GIJÓN`).
   - Cobertura completa de nombres de residencia en `Select Case ActiveSheet.Name` (`"RESIDENCIA GIJÓN", "RESIDENCIA GIJON", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"`).
   - Codificación ANSI nativa Windows-1252 sin cabeceras UTF-8 BOM.

3. **Carátula Pantalla Completa Casa Cuartel Antigua en Hoja INICIO (`ThisWorkbook.cls`)**:
   - Integración de la fotografía histórica `Foto Casa Cuartel antigua.png` como imagen principal de bienvenida.
   - Reescalado automático en `AjustarCaratulaInicio`: asignación precisa a `Application.UsableWidth` y `Application.UsableHeight` con deformación leve controlada (`shp.LockAspectRatio = msoFalse`), centrado en origen (`Left = 0, Top = 0`) y colocación en el fondo (`shp.ZOrder msoSendToBack`).

4. **Rediseño Compacto y Ergonómico del Cuadro de Identificación (`frmLogin.frm`)**:
   - Reducción del tamaño general de la tarjeta de identificación en aprox. un 30% ($170\text{ pt}$ ancho $\times 130\text{ pt}$ alto) con fondo blanco limpio.
   - Reducción del tamaño de los campos de texto `txtUsuario` y `txtPassword` en un 20% ($110\text{ pt}$ ancho $\times 16\text{ pt}$ alto).
   - Reubicación vertical de los botones `btnEntrar` y `btnSalir`, elevándolos hacia la casilla de Clave para prevenir recortes en el borde inferior.
   - Anclaje milimétrico entre los dos Guardias Civiles: centrado horizontal y encaje perfecto en el patio empedrado, dejando a ambos guardias totalmente visibles y sin solapamientos.

5. **Erradicación Total de Errores de Sintaxis por Continuaciones de Línea (` _`)**:
   - Diagnóstico profundo a nivel de byte: supresión de saltos dobles `\r\r\n` que el analizador léxico de VBA interpretaba como líneas vacías intermedias entre sentencias concatenadas (`MsgBox "No se puede conectar a la base de datos Access." & vbCrLf & vbCrLf & _`).
   - Saneamiento de todas las continuaciones multilínea en `ThisWorkbook.cls` (18 ocurrencias) y `EvitarDuplicidadSolicitudesGyS.bas` (11 ocurrencias).
   - Validación activa en Excel COM (`FindControl(1, 578).Enabled == False`), certificando **0 errores de compilación** en todo el proyecto VBA.

6. **Gobernanza Técnica: Protocolo Estricto de Codificación y Cero Mojibake en VBA**:
   - Codificación ANSI Windows-1252 nativa con 0 caracteres corruptos ni secuencias UTF-8 multibyte.
   - Regla estricta de nombres literal `GIJÓN` (0 usos de `Chr(211)`).
   - Validación satisfactoria por el Guardián Pre-Commit de Git.

7. **Preservación Total e Incondicional de los 171 Componentes VBA**:
   - Tarifas dinámicas en facturas PAX vinculadas a `TARIFAS UNIDAD`.
   - Saneamiento de hojas LOG y blindaje contra inyección de fórmulas (`'=`).
   - Sincronización relacional inversa Top 500 y desambiguación de login por sede.
   - Sesión persistente auto-reparable en Nombres Definidos ocultos.
   - Supresión del reloj de arena en columna T de `Hoja2 (RESIDENCIA GIJÓN)`.
   - Diálogo interactivo y salida suave en `Workbook_BeforeClose`.

---

## 2. Histórico de Versiones de Referencia Archivadas

| Hito Temporal | Libro XLSM en `Historico\` | Base de Datos ACCDB en `Historico\` | MD5 Criptográfico XLSM | Tamaño XLSM |
| :---: | :---| :---| :---: | :---: |
| **2026-10-22 18:55** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-22_1855.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-22_1855.accdb` | `275253EADF79F252644FFBF9140516A3` | 6.748.351 bytes |
| **2026-10-21 23:30** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-21_2330.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-21_2330.accdb` | `2507104C574B3DD48E11DD9A0FD76A76` | 7.188.546 bytes |
| **2026-10-21 00:55** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-21_0055.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-21_0055.accdb` | `E8AE2E58DC520C55042E93C8992945F9` | 7.171.197 bytes |
| **2026-10-20 23:15** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_2315.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_2315.accdb` | `5A23A4AB3A34A88442E6C28D5B1FB021` | 7.151.821 bytes |
| **2026-10-20 19:35** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_1935.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_1935.accdb` | `1B4AC09D12DF1BEFA8520EBEC0A0F811` | 7.152.489 bytes |
| **2026-10-20 01:00** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_0100.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_0100.accdb` | `C8F9A4B7D26E4A0C7B8E9D1F3A5C6E7F` | 7.133.373 bytes |
| **2026-10-19 21:35** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_2135.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_2135.accdb` | `6F22BD002241F740BAE7F280D597A2E9` | 6.429.438 bytes |
| **2026-10-19 01:50** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_0150.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_0150.accdb` | `B40EF0EE044439C904AE4BEA973641FA` | 6.426.870 bytes |
| **2026-10-18 21:25** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_2125.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_2125.accdb` | `A09B66046CF22A4E002C9FA8F0BE496A` | 6.314.767 bytes |
| **2026-10-18 03:20** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_0320.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_0320.accdb` | `F30349A731114ED4D55140880C199026` | 5.792.751 bytes |
| **2026-10-17 23:15** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-17_2315.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-17_2315.accdb` | `C3D6AA392BF93EC0F3C8FF9FE29F3424` | 5.769.867 bytes |
| **2026-10-11 02:58** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-11.xlsm` | — | `63F69EBE492938B12A8DE5F4B650D799` | 6.530.556 bytes |

---

## 3. Matriz de Componentes Clave en este Hito

| Componente | Tipo | Función Principal en Hito 2026-10-22 19:45 |
| :--- | :--- | :--- |
| [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) | Objeto Excel | Auto-Fit a pantalla completa de carátula en `INICIO`, reposición dinámica de barras de desplazamiento en `Workbook_Open` y `Workbook_SheetActivate`, erradicación de continuaciones rotas ` _`, confirmación interactiva de salida y protección de LOGs. |
| [frmLogin.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmLogin.frm) | UserForm | Tarjeta de identificación compactada al 70%, campos reducidos al 80%, botones elevados y anclaje milimétrico en el patio empedrado entre guardias civiles. |
| [BotonIzdaDchaHojasRESIDENCIAS.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonIzdaDchaHojasRESIDENCIAS.bas) | Módulo Estándar | Navegación por bloques de columnas (GoToPrevious / GoToNext) con saneamiento ortográfico Windows-1252 y soporte íntegro para GIJÓN, SOTO y OVIEDO. |
| [EvitarDuplicidadSolicitudesGyS.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/EvitarDuplicidadSolicitudesGyS.bas) | Módulo Estándar | Saneamiento de continuaciones rotas ` _`, sustitución de flechas por `->`, acentuación ortográfica completa y consulta relacional a BD Access. |
| [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) | Módulo Estándar | Núcleo ADODB multiusuario: sesión persistente en Nombres Definidos ocultos, sincronización relacional Top 500, UPSERT y validación de conexión. |
| [Hoja2 (RESIDENCIA GIJÓN)](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja2_Gijon.cls) | Hoja Excel | Supresión del reloj de arena en columna T (*DTO. FAM. NUM.*) y validación de coherencia habitacional en columna AB (*Grabación Solicitud*). |
