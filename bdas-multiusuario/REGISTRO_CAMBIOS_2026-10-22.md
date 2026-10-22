# REGISTRO DE CAMBIOS - 2026-10-22

## Hito 2026-10-22 19:35: Consolidación y Promoción a Versión de Referencia Oficial

### 1. Resumen Ejecutivo
Se ha promocionado formalmente la versión activa de trabajo `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando el hito previo (2026-10-21 23:30) en el histórico, actualizando el script de restauración inmediata (`Restaurar-Version-Referencia.ps1`), registrando exhaustivamente la trazabilidad de copias de seguridad incrementales diarias y sincronizando la copia de despliegue en red (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`).

Asimismo, se ha implementado el **restablecimiento dinámico e integral de las barras de desplazamiento (vertical y horizontal)** en todas las hojas operativas de la aplicación una vez superado el proceso de identificación inicial, garantizando la navegación completa por todas las columnas y filas de datos (especialmente en `RESIDENCIA OVIEDO`, `RESIDENCIA GIJÓN` y `RESIDENCIA SOTO`).

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Atributos |
| :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 6.755.115 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | Solo Lectura (`ReadOnly = True`) |
| **Versión de Referencia Oficial (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Solo Lectura (`ReadOnly = True`) |
| **Versión Activa de Trabajo (App)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.755.115 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 6.755.115 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | Lectura / Escritura |
| **Base de Datos Activa (Backend)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Lectura / Escritura |
| **Copia Incremental Post Barras XLSM** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-22_1935_POST_BARRAS_SCROLL.xlsm` | 6.755.115 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | Respaldo Incremental |
| **Copia Incremental Pre Barras XLSM** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-22_1925_PRE_BARRAS_SCROLL.xlsm` | 6.748.351 | `275253EADF79F252644FFBF9140516A3` | Respaldo Incremental |
| **Copia Incremental Ref. ACCDB** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.accdb` | 6.729.728 | `133BA235E1D42C42AF2448890E3ED64B` | Respaldo Incremental |
| **Histórico Referencia Previa (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-21_2330.xlsm` | 7.188.546 | `2507104C574B3DD48E11DD9A0FD76A76` | Solo Lectura (`ReadOnly = True`) |
| **Histórico Referencia Previa (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-21_2330.accdb` | 6.725.632 | `D51AE318D49BC74C4BC738F5B58B2B51` | Solo Lectura (`ReadOnly = True`) |

---

### 3. Vinculación de Copias Incrementales con Funcionalidades y Código

Durante la sesión técnica del 22/10/2026 se generaron 6 copias de seguridad incrementales en `H:\ResidenciaBD\BackupsDiarios\`, vinculadas de forma biunívoca a cada funcionalidad desarrollada:

| Copia Incremental XLSM / ACCDB | Fecha / Hora | MD5 Criptográfico | Tamaño (bytes) | Funcionalidad y Cambios Asociados |
| :---| :---: | :---:| :---:| :---|
| `BDAS_Multiusuario_BACKUP_2026-10-22_1755.xlsm` | 22/10/2026 17:42 | `013559F77BF54B057A763E5F72365EC1` | 7.203.037 | **Rediseño Compacto de Tarjeta Login y Foto Casa Cuartel Antigua:** Incorporación de la fotografía histórica `Foto Casa Cuartel antigua.png` en la carátula de inicio. Rediseño del modal `frmLogin` con reducción del 30% en tamaño general (de 240x165 a 170x130 pt), reducción del 20% en campos de texto Usuario y Clave (110x16 pt) y elevación de los botones ENTRAR/SALIR hacia el campo de Clave, anclando la tarjeta herméticamente en el patio empedrado entre los dos guardias civiles sin recortes. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1815.xlsm` | 22/10/2026 18:01 | `A8259062895A56600B67A7E75C5C8DE9` | 7.214.915 | **Ajuste a Pantalla Completa de Carátula y Bloqueo de Recorrido de Hoja INICIO:** Actualización de `AjustarCaratulaInicio` en `ThisWorkbook.cls` (`shp.LockAspectRatio = msoFalse`, asignación a `UsableWidth` y `UsableHeight`, `shp.Left = 0, Top = 0`, `shp.ZOrder msoSendToBack`). Ocultación total de barras de desplazamiento vertical y horizontal en la carátula (`DisplayVerticalScrollBar = False`, `DisplayHorizontalScrollBar = False`) y congelación hermética en celda única con `ScrollArea = "A1:A1"`. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1845.xlsm` | 22/10/2026 18:41 | `275253EADF79F252644FFBF9140516A3` | 6.748.351 | **Erradicación del Error de Sintaxis por Continuaciones Rotas ` _` y Saneamiento CRLF Puro:** Diagnóstico de causa raíz por generación de saltos `\r\r\n` que insertaban líneas en blanco tras el guion bajo ` _` en instrucciones multilínea (`MsgBox "No se puede conectar a la base de datos Access." & vbCrLf & vbCrLf & _`). Saneamiento de 18 continuaciones rotas en `ThisWorkbook.cls` y 11 en `EvitarDuplicidadSolicitudesGyS.bas`. Normalización a codificación Windows-1252 ANSI nativa con 0 secuencias UTF-8 multibyte (`\xc3`), preservación estricta de `GIJÓN` literal (0 `Chr(211)`). Compilación limpia certificada en Excel COM (`FindControl(1, 578).Enabled == False`). |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.xlsm` | 22/10/2026 18:55 | `275253EADF79F252644FFBF9140516A3` | 6.748.351 | **Promoción Inicial a Versión de Referencia:** Consolidación formal de la versión de trabajo, establecimiento de réplicas maestras en `H:\ResidenciaBD\Plantilla\VersionReferencia\`, archivado histórico del hito 2026-10-21 23:30 y actualización del script de restauración. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1925_PRE_BARRAS_SCROLL.xlsm` | 22/10/2026 19:25 | `275253EADF79F252644FFBF9140516A3` | 6.748.351 | **Punto de Control Pre-Restauración de Barras de Scroll:** Salvaguarda del estado previo a la reactivación de las barras de desplazamiento operativas. |
| `BDAS_Multiusuario_BACKUP_2026-10-22_1935_POST_BARRAS_SCROLL.xlsm` | 22/10/2026 19:35 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | 6.755.115 | **Restauración Integral de Barras de Desplazamiento y Promoción Definitiva:** Restablecimiento automático de `DisplayVerticalScrollBar = True` y `DisplayHorizontalScrollBar = True` (junto con encabezados y cuadrículas) al finalizar el arranque en `Workbook_Open` y de forma reactiva en `Workbook_SheetActivate` para cualquier hoja que no sea `INICIO`. Saneamiento de `BotonIzdaDchaHojasRESIDENCIAS.bas` en Windows-1252 ANSI puro. Verificación COM activa con 0 errores de compilación y consolidación de la réplica de referencia. |
| `Residencia_BE_BACKUP_2026-10-22_1855_VERSION_REFERENCIA.accdb` | 22/10/2026 18:55 | `133BA235E1D42C42AF2448890E3ED64B` | 6.729.728 | **Respaldo Incremental de Base de Datos Backend:** Copia de seguridad completa de la base de datos Access backend sincronizada con el hito de referencia 2026-10-22. |

---

### 4. Alcance Técnico Consolidado en este Hito

1. **Restauración Dinámica e Integral de las Barras de Desplazamiento (`ThisWorkbook.cls`)**:
   - **Comportamiento en Carátula de Inicio (`INICIO`):** Las barras de desplazamiento vertical y horizontal se mantienen ocultas (`DisplayVerticalScrollBar = False`, `DisplayHorizontalScrollBar = False`) junto con los encabezados y cuadrículas exclusivamente durante la pantalla de presentación e introducción de credenciales, con la celda confinada a `ScrollArea = "A1:A1"`.
   - **Comportamiento en Carga Completa (`Workbook_Open`):** Una vez completado el login y sincronizadas las hojas desde Access, inmediatamente tras el Paso 8, se reponen de forma explícita las barras en la ventana activa (`ActiveWindow.DisplayVerticalScrollBar = True`, `ActiveWindow.DisplayHorizontalScrollBar = True`, `DisplayHeadings = True`, `DisplayGridlines = True`).
   - **Gobernanza Reactiva en Cambio de Hoja (`Workbook_SheetActivate`):** Se intercepta el evento de activación de cualquier hoja: si se activa `INICIO`, se ocultan las barras; para cualquier otra hoja del libro (`RESIDENCIA OVIEDO`, `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO`, `CALENDARIO ...`, `RESUMEN ...`, `LOG ...`, etc.), se reactivan automáticamente ambas barras de scroll, asegurando la navegación completa por todas las columnas (más allá de la columna S y AE) y filas.
   - **Cierre Limpio en Contingencia (`CleanClose`):** Si el login es cancelado o falla, se restauran las barras antes de cerrar para no dejar la aplicación Excel en un estado degradado.

2. **Saneamiento de Módulo de Navegación por Columnas (`BotonIzdaDchaHojasRESIDENCIAS.bas`)**:
   - Supresión de caracteres corruptos mojibake (`GIJï¿½N` $\to$ `GIJÓN`).
   - Cobertura completa de nombres de residencia en `Select Case ActiveSheet.Name` (`"RESIDENCIA GIJÓN", "RESIDENCIA GIJON", "RESIDENCIA SOTO", "RESIDENCIA OVIEDO"`).
   - Conversión a codificación ANSI nativa Windows-1252 sin cabeceras UTF-8 BOM.

3. **Carátula Pantalla Completa Casa Cuartel Antigua en Hoja INICIO (`ThisWorkbook.cls`)**:
   - Reemplazo y ajuste de la fotografía de portada por la imagen histórica (`Foto Casa Cuartel antigua.png`).
   - Reescalado automático en `AjustarCaratulaInicio`: asignación precisa a `Application.UsableWidth` y `Application.UsableHeight` con deformación leve controlada (`shp.LockAspectRatio = msoFalse`), centrado en origen (`Left = 0, Top = 0`) y envío al fondo (`shp.ZOrder msoSendToBack`).

4. **Rediseño Compacto y Ergonómico del Cuadro de Identificación (`frmLogin.frm`)**:
   - Reducción general de la tarjeta en ~30% ($170\times 130\text{ pt}$) con fondo blanco limpio.
   - Reducción de campos Usuario y Clave en un 20% ($110\times 16\text{ pt}$) y elevación de botones ENTRAR/SALIR.
   - Anclaje milimétrico en el patio empedrado entre los dos guardias civiles sin recortes inferiores ni superposiciones.

5. **Erradicación Definitiva de Errores de Sintaxis por Continuaciones de Línea (` _`)**:
   - Saneamiento de todas las continuaciones multilínea en `ThisWorkbook.cls` y `EvitarDuplicidadSolicitudesGyS.bas`.
   - Verificación de compilación en Excel COM (`FindControl(1, 578).Enabled == False`), certificando **0 errores de compilación**.

6. **Gobernanza Técnica Anti-Mojibake y Calidad ECRMC**:
   - Todos los componentes grabados en Windows-1252 (ANSI nativo).
   - Cumplimiento incondicional de la regla estricta de nombres `GIJÓN` (0 usos de `Chr(211)`).
   - Superación del Control de Puerta Pre-Entrega y validación satisfactoria por el hook pre-commit de Git.
