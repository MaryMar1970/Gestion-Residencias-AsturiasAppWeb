# REGISTRO DE CAMBIOS - 2026-10-22

## Hito 2026-10-22 18:55: Consolidación y Promoción a Versión de Referencia Oficial

### 1. Resumen Ejecutivo
Se ha promocionado formalmente la versión activa de trabajo `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando el hito previo (2026-10-21 23:30) en el histórico, actualizando el script de restauración inmediata (`Restaurar-Version-Referencia.ps1`), registrando exhaustivamente la trazabilidad de copias de seguridad incrementales diarias y sincronizando la copia de despliegue en red (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`).

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Atributos |
| :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 6.748.351 | `275253EADF79F252644FFBF9140516A3` | Solo Lectura (`ReadOnly = True`) |
| **Versión de Referencia Oficial (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Solo Lectura (`ReadOnly = True`) |
| **Versión Activa de Trabajo (App)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.748.351 | `275253EADF79F252644FFBF9140516A3` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 6.748.351 | `275253EADF79F252644FFBF9140516A3` | Lectura / Escritura |
| **Base de Datos Activa (Backend)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Lectura / Escritura |
| **Copia Incremental Ref. XLSM** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.xlsm` | 6.748.351 | `275253EADF79F252644FFBF9140516A3` | Respaldo |
| **Copia Incremental Ref. ACCDB** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Respaldo |
| **Histórico Referencia Previa (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-21_2330.xlsm` | 7.188.546 | `2507104C574B3DD48E11DD9A0FD76A76` | Solo Lectura (`ReadOnly = True`) |
| **Histórico Referencia Previa (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-21_2330.accdb` | 6.725.632 | `D51AE318D49BC74C4BC738F5B58B2B51` | Solo Lectura (`ReadOnly = True`) |

---

### 3. Vinculación de Copias Incrementales con Funcionalidades y Código

Durante la sesión técnica del 22/10/2026 se generaron 5 copias de seguridad incrementales en `H:\ResidenciaBD\BackupsDiarios\`, vinculadas de forma biunívoca a cada funcionalidad desarrollada:

| Copia Incremental XLSM / ACCDB | Fecha / Hora | MD5 Criptográfico | Tamaño (bytes) | Funcionalidad y Cambios Asociados |
| :---| :---: | :---:| :---:| :---|
| `BDAS_Multiusuario_BACKUP_2026-10-22_1755.xlsm` | 22/10/2026 17:42 | `013559F77BF54B057A763E5F72365EC1` | 7.203.037 | **Rediseño Compacto de Tarjeta Login y Foto Casa Cuartel Antigua:** Incorporación de la fotografía histórica `Foto Casa Cuartel antigua.png` en la carátula de inicio. Rediseño del modal `frmLogin` con reducción del 30% en tamaño general (de 240x165 a 170x130 pt), reducción del 20% en campos de texto Usuario y Clave (110x16 pt) y elevación de los botones ENTRAR/SALIR hacia el campo de Clave, anclando la tarjeta herméticamente en el patio empedrado entre los dos guardias civiles sin recortes. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1815.xlsm` | 22/10/2026 18:01 | `A8259062895A56600B67A7E75C5C8DE9` | 7.214.915 | **Ajuste a Pantalla Completa de Carátula y Bloqueo de Recorrido de Hoja INICIO:** Actualización de `AjustarCaratulaInicio` en `ThisWorkbook.cls` (`shp.LockAspectRatio = msoFalse`, asignación a `UsableWidth` y `UsableHeight`, `shp.Left = 0, Top = 0`, `shp.ZOrder msoSendToBack`). Ocultación total de barras de desplazamiento vertical y horizontal en la carátula (`DisplayVerticalScrollBar = False`, `DisplayHorizontalScrollBar = False`) y congelación hermética en celda única con `ScrollArea = "A1:A1"`. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1845.xlsm` | 22/10/2026 18:41 | `275253EADF79F252644FFBF9140516A3` | 6.748.351 | **Erradicación del Error de Sintaxis por Continuaciones Rotas ` _` y Saneamiento CRLF Puro:** Diagnóstico de causa raíz por generación de saltos `\r\r\n` que insertaban líneas en blanco tras el guion bajo ` _` en instrucciones multilínea (`MsgBox "No se puede conectar a la base de datos Access." & vbCrLf & vbCrLf & _`). Saneamiento de 18 continuaciones rotas en `ThisWorkbook.cls` y 11 en `EvitarDuplicidadSolicitudesGyS.bas`. Normalización a codificación Windows-1252 ANSI nativa con 0 secuencias UTF-8 multibyte (`\xc3`), preservación estricta de `GIJÓN` literal (0 `Chr(211)`). Compilación limpia certificada en Excel COM (`FindControl(1, 578).Enabled == False`). |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.xlsm` | 22/10/2026 18:55 | `275253EADF79F252644FFBF9140516A3` | 6.748.351 | **Promoción Oficial a Versión de Referencia:** Consolidación formal de la versión de trabajo, establecimiento de réplicas maestras en `H:\ResidenciaBD\Plantilla\VersionReferencia\`, archivado histórico del hito 2026-10-21 23:30 y actualización del script de restauración. |
| `Residencia_BE_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.accdb` | 22/10/2026 18:55 | `133BA235E1D42C42AF2448890E3ED64B` | 6.729.728 | **Respaldo Incremental de Base de Datos Backend:** Copia de seguridad completa de la base de datos Access backend sincronizada con el hito de referencia 2026-10-22 18:55. |

---

### 4. Alcance Técnico Consolidado en este Hito

1. **Carátula Pantalla Completa Casa Cuartel Antigua en Hoja INICIO (`ThisWorkbook.cls`)**:
   - Reemplazo y ajuste de la fotografía de portada por la imagen histórica (`Foto Casa Cuartel antigua.png`).
   - Reescalado automático en `AjustarCaratulaInicio`: asignación precisa a `Application.UsableWidth` y `Application.UsableHeight` con deformación leve controlada (`shp.LockAspectRatio = msoFalse`), centrado en origen (`Left = 0, Top = 0`) y envío al fondo (`shp.ZOrder msoSendToBack`).
   - Ocultación garantizada de barras de desplazamiento vertical y horizontal en la carátula (`DisplayVerticalScrollBar = False`, `DisplayHorizontalScrollBar = False`) y confinamiento del cursor a la celda `A1` mediante `ScrollArea = "A1:A1"` para eliminar cualquier desplazamiento o recorrido visual indeseado.

2. **Rediseño Compacto y Ergonómico del Cuadro de Identificación (`frmLogin.frm`)**:
   - Tarjeta estilizada y proporcionada: reducción en aprox. un 30% respecto al diseño anterior, pasando a dimensiones óptimas de $170\text{ pt}$ ancho $\times 130\text{ pt}$ alto con fondo blanco limpio.
   - Reducción del tamaño de los campos de texto `txtUsuario` y `txtPassword` en un 20% ($110\text{ pt}$ ancho $\times 16\text{ pt}$ alto), aumentando la ligereza visual.
   - Reubicación vertical de los botones `btnEntrar` y `btnSalir`, elevándolos hacia la casilla de Clave para prevenir recortes en el borde inferior.
   - Posicionamiento milimétrico calculado: la tarjeta queda enmarcada limpiamente en el patio empedrado entre los dos guardias civiles, sin tapar sus figuras ni desbordar la composición.
   - Flujo por teclado intacto: foco inicial automático en `txtUsuario`, avance con `Enter` a `txtPassword`, y login instantáneo con `Enter`.

3. **Erradicación Definitiva de Errores de Sintaxis por Continuaciones de Línea (` _`)**:
   - Diagnóstico profundo a nivel de byte: resolución de discrepancias por saltos de línea dobles (`\r\r\n`) que el analizador léxico de VBA interpretaba como líneas vacías entre sentencias concatenadas.
   - Saneamiento de todas las continuaciones multilínea en [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) (18 ocurrencias corregidas) y [EvitarDuplicidadSolicitudesGyS.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/EvitarDuplicidadSolicitudesGyS.bas) (11 ocurrencias corregidas).
   - Verificación de compilación en Excel COM (`FindControl(1, 578).Enabled == False`), certificando **0 errores de compilación** en todo el proyecto VBA.

4. **Gobernanza Técnica Anti-Mojibake y Calidad ECRMC**:
   - Todos los componentes grabados en Windows-1252 (ANSI nativo), erradicando secuencias multibyte UTF-8 corruptas.
   - Cumplimiento incondicional de la regla estricta de nombres `GIJÓN` (0 usos de `Chr(211)`).
   - Superación del Control de Puerta Pre-Entrega y validación satisfactoria por el hook pre-commit de Git.

5. **Preservación Total de los 171 Componentes Previos**:
   - Tarifas dinámicas en facturas PAX vinculadas a `TARIFAS UNIDAD`.
   - Normalización de hojas LOG y blindaje contra inyección de fórmulas (`'=`).
   - Sincronización relacional Top 500 y desambiguación de login por sede.
   - Sesión persistente auto-reparable en Nombres Definidos ocultos.
   - Supresión del reloj de arena en columna T de `Hoja2 (RESIDENCIA GIJÓN)`.
   - Diálogo interactivo y salida suave en `Workbook_BeforeClose`.
