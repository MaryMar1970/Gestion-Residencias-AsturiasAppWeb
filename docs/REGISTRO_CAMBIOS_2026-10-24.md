# REGISTRO DE CAMBIOS - 2026-10-24

## Hito 2026-10-24 00:45: Consolidación y Promoción a Versión de Referencia Oficial (Fases 1 y 2 Completadas)

### 1. Resumen Ejecutivo
Se ha promocionado formalmente la versión activa de trabajo `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando los hitos previos en el histórico (`Historico\`), actualizando los scripts de promoción y restauración inmediata (`Promocionar-Version-Referencia.ps1`, `Restaurar-Version-Referencia.ps1`), registrando exhaustivamente la trazabilidad de copias de seguridad incrementales diarias vinculadas al código y sincronizando la copia de despliegue en red (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`).

Este hito consolida de forma definitiva dos grandes avances técnicos auditados por el Arquitecto Principal de Software y validados localmente por el usuario con **0 errores de compilación**:
1. **Fase 1 (Auditoría Integral y Refactorización Arquitectónica - REF-01 a REF-16):** Centralización de consultas SQL, blindaje de variables globales, optimización contable de facturas PAX, eliminación de duplicidades y saneamiento ortográfico global en 10 módulos y hojas.
2. **Fase 2 (Modernización UX/UI Fluent Card y Anti-Parpadeo - Secciones G.1 a G.5 Completas):** Eliminación total del parpadeo por bandas en el arranque mediante consolidación de buffer gráfico e idempotencia geométrica, pre-carga en memoria y amortiguación de repintado en `modDatabase.LoginUsuario`, rediseño visual Fluent Card en `frmLogin` y `frmSelectorResidencia`, y modernización contemporánea de la notificación flotante (Toast no modal) `frmAviso`.

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Hash SHA-256 | Atributos |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 6.838.715 | `C61E591EDD8BCAB9D789949D471FF617` | `FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6` | Solo Lectura (`ReadOnly = True`) |
| **Versión de Referencia Oficial (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.733.824 | `A5B86926058CCF73DEDA7F0C8F83A323` | `DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89` | Solo Lectura (`ReadOnly = True`) |
| **Versión Activa de Trabajo (App)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.838.715 | `C61E591EDD8BCAB9D789949D471FF617` | `FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 6.838.715 | `C61E591EDD8BCAB9D789949D471FF617` | `FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6` | Lectura / Escritura |
| **Base de Datos Activa (Backend)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.733.824 | `A5B86926058CCF73DEDA7F0C8F83A323` | `DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89` | Lectura / Escritura |
| **Copia Cierre Ref. XLSM (00:45)** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-24_0045_VERSION_REFERENCIA.xlsm` | 6.838.715 | `C61E591EDD8BCAB9D789949D471FF617` | `FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6` | Respaldo Incremental |
| **Copia Cierre Ref. ACCDB (00:45)** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-24_0045_VERSION_REFERENCIA.accdb` | 6.733.824 | `A5B86926058CCF73DEDA7F0C8F83A323` | `DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89` | Respaldo Incremental |
| **Copia Pre-Adenda XLSM (23:30)** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-23_2330_PRE_ADENDA_FASE2_G2_G5.xlsm` | 6.791.236 | `F1FAEFDEE6953D431FBECEB6F775B3D5` | `2DD40A56047D22BA47F9FAA8EAF78D4C843448F30A33A29968E059E8D39339C3` | Respaldo Incremental |
| **Copia Pre-Adenda ACCDB (23:30)** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-23_2330_PRE_ADENDA_FASE2_G2_G5.accdb` | 6.733.824 | `A5B86926058CCF73DEDA7F0C8F83A323` | `2F13F77CC340D0C92EDA3286BFB144D8E086F2738396116C174771DE09D43EA5` | Respaldo Incremental |
| **Copia Pre-Fase 2 XLSM (21:45)** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-23_2145_PRE_FASE2_FLUENT_UI.xlsm` | 6.776.283 | `3C56F3F34B22B57B2FDF5D26C2A7B543` | `87A1D8D76F95E18452858921FF9DD61E2AB1EF3638C33534D3BF159BF72A162D` | Respaldo Incremental |
| **Copia Pre-Fase 2 ACCDB (21:45)** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-23_2145_PRE_FASE2_FLUENT_UI.accdb` | 6.733.824 | `A5B86926058CCF73DEDA7F0C8F83A323` | `DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89` | Respaldo Incremental |
| **Histórico Referencia 22-Oct (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-22_1945.xlsm` | 6.755.115 | `DB8FFA1648DD2917C385E2AAED7E9BD8` | `E3BDC5D99B5897F2997C6F12C889410EDEFBDD91F0BE71249969F218BDE12185` | Solo Lectura (`ReadOnly = True`) |
| **Histórico Referencia 22-Oct (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-22_1945.accdb` | 6.729.728 | `6E41A8521D884B740907CD9DEF187CEA` | `FA72BE0F51B9DAEE57EF9D760A8149845348F371493A60EEF720A8E3F45E6746` | Solo Lectura (`ReadOnly = True`) |

---

### 3. Registro de Funcionalidades y Mejoras Consolidadas en el Código

#### A. Modernización UX/UI Fluent Card y Anti-Parpadeo (Secciones G.1 a G.5)
1. **G.1.1 — Supresión Total de Eventos y Consolidación de Buffer (`ThisWorkbook.cls` / `Workbook_Open`):**
   - Se fija la vista en la hoja `INICIO` antes de iniciar cualquier proceso de renderizado.
   - Consolidación del buffer gráfico de Office 2016 con `ActiveWindow.SmallScroll Down:=0` antes de invocar el cuadro de login, eliminando las bandas blancas y repintados parciales.
2. **G.1.2 — Idempotencia Geométrica en Carátula (`ThisWorkbook.cls` / `AjustarCaratulaInicio`):**
   - Incorporación de umbral de tolerancia ($\pm 2	ext{ pt}$): si las dimensiones de la imagen `ImgCaratulaCasaCuartel` coinciden con el área visible de la ventana, se omite el proceso `Unprotect/Resize/Protect`.
   - Evita invalidar la caché de memoria GDI de la imagen de la Casa Cuartel Antigua ante llamadas redundantes o cambios de foco.
3. **G.2 (Revisión 2) — Transición Amortiguada y Pre-carga en Memoria (`modDatabase.bas` / `LoginUsuario`):**
   - Pre-carga del formulario de login (`Load frmLogin`) en memoria antes de entrar al bucle modal.
   - Congelación de repintado (`Application.screenUpdating = False`) antes de `Unload frmLogin` para evitar destellos sobre la carátula.
   - Pre-carga del selector multisede (`Load frmSelectorResidencia`) reactivando pantalla en el instante exacto de mostrarlo (`Show vbModal`).
   - Salidas anticipadas unificadas hacia la etiqueta `GoTo CleanUp` para garantizar la descarga limpia de ambos formularios.
   - Preservación íntegra de la seguridad de intentos (`MAX_INTENTOS_LOGIN = 3`), variables de sesión (`m_NombreUsuario`, `m_ResidenciaActiva`), activación de hojas (`MostrarHojasResidencia`), refresco del Ribbon (`InicioGestionResidencias.RefrescarRibbon`), registro en Access (`InsertarLog`), registro en Excel (`ModuloLOG.RegistrarLoginLOG`), protección de logs y sesión persistente (`GuardarSesionEnLibro`).
4. **G.3.1 y G.3.2 — Tarjeta Fluent Card e Interacción Dinámica (`frmLogin.frm`):**
   - Superficie Fluent Card blanca pura (`RGB(255, 255, 255)`), franja superior de acento institucional verde Guardia Civil (`RGB(27, 67, 50)` de $3.5	ext{ pt}$).
   - Campos planos (*Flat Inputs*) con borde sutil en reposo (`RGB(203, 213, 225)`) y borde verde corporativo dinámico al recibir foco.
   - Tipografía moderna *Segoe UI* en todos los controles y sustitución del asterisco de clave por el punto contemporáneo `ChrW(9679)`.
   - Sincronización de tecla `Enter` en `txtUsuario_KeyDown` consumiendo la pulsación (`KeyCode = 0`) para evitar pitido del sistema y transferir el foco a `txtClave`.
   - Micro-interacciones hover en botones y supresión del reposicionamiento en `UserForm_Activate` (erradicando el salto y parpadeo de bandas).
5. **G.4.1 — Selector Multisede Fluent Card (`frmSelectorResidencia.frm`):**
   - Rediseño con cabecera verde institucional, fondo blanco y tipografía *Segoe UI*.
   - Contenedor plano de selección con casillas de verificación nativas.
   - Botón de acción primario y normalización estricta del literal `GIJÓN` sin funciones de escape prohibidas.
6. **G.5 (Revisión 2) — Notificación Flotante Toast No Modal (`frmAviso.frm`):**
   - Fondo blanco `#FFFFFF`, franja superior verde de $4	ext{ pt}$ y tipografía *Segoe UI* en negrita sobre el control `Me.lblAviso`.
   - Preservación estricta de su carácter no modal (`Show vbModeless`), evitando bloqueos en la navegación de celdas.
   - Desprogramación defensiva del temporizador `Application.OnTime` si el usuario selecciona rápidamente varias celdas consecutivas de teléfono o DNI en la columna 28 (`AB`) de las residencias.
   - Preservación íntegra de los métodos públicos `MostrarAviso(texto, segundos)` y `CerrarAviso()`.

#### B. Refactorización Arquitectónica y Auditoría Integral (Fase 1: REF-01 a REF-16)
1. **Centralización y Persistencia de Base de Datos (`modDatabase.bas`):**
   - Centralización de sentencias SQL y eliminación de consultas duplicadas.
   - Robustecimiento de transacciones y optimización de reconexiones automáticas ADODB.
2. **Saneamiento Integral de Auditoría (`ModuloLOG.bas`):**
   - Parseo estricto por columnas: `[Usuario, Fecha/Hora, Nº Orden, Columna, Valor anterior, Valor nuevo]`.
   - Protección contra inyección de fórmulas (apóstrofo defensivo ante cadenas que inician con `'='`).
   - Sincronización relacional de los 500 registros más recientes de Access sin truncamiento histórico.
3. **Módulos y Hojas Operativas de Residencia:**
   - Saneamiento en `Hoja2_RESIDENCIA GIJÓN.cls`, `Hoja20_RESIDENCIA SOTO.cls` y `Hoja24_RESIDENCIA OVIEDO.cls`.
   - Erradicación de retardos por eventos redundantes en selección de celdas.
   - Cuadre contable exacto al céntimo en facturas PAX vinculadas a `TARIFAS UNIDAD`.
   - Corrección en `AsignarNumFactura.bas`, `MarcarSiPagadosEnResidencia.bas`, `ReevaluacionSolicitudes.bas`, `Hoja9_RESUMEN SOTO.cls` y `Hoja7_RESUMEN GIJÓN.cls`.

---

### 4. Gobernanza Anti-Mojibake y Control de Puerta
- **Codificación:** 100% nativa **Windows-1252 (ANSI)** con terminaciones **CRLF (`\r\n`)**.
- **Control de Caracteres Corruptos:** 0 patrones de mojibake (`Ã`, `Â`, `ǭ`, `âž`, `ï»¿`) en los 171 componentes.
- **Regla Estricta `GIJÓN`:** 0 llamadas a `Chr(211)` o concatenaciones `GIJ" & Chr(211) & "N`.
- **Compilación en Excel COM:** Verificada al 100% (`VBE.CommandBars.FindControl(1, 578).Enabled == False`).

---

## Hito 2026-10-24 00:55: Promoción Expresa a GOLDEN MASTER Oficial

### 1. Resumen Ejecutivo
Atendiendo a la **petición expresa e inequívoca del usuario**, se ha promocionado la versión activa de trabajo BDAS_Multiusuario.xlsm a la máxima categoría de salvaguarda del sistema: **GOLDEN MASTER OFICIAL**.

Se ha preservado el Golden Master fundacional previo (2026-09-18) en el histórico (H:\ResidenciaBD\Plantilla\GoldenMaster\Historico\BDAS_Multiusuario_GOLDEN_MASTER_2026-09-18.xlsm), se ha establecido la nueva réplica protegida en modo Solo Lectura en H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm y Residencia_BE_GOLDEN_MASTER.accdb, se ha generado la etiqueta en git GOLDEN_MASTER_2026-10-24, y se ha actualizado el script [Restaurar-Golden-Master.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Golden-Master.ps1) para garantizar la restauración instantánea en caso de desastre.

### 2. Identificadores Criptográficos del Nuevo Golden Master
- **Archivo XLSM:** H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm
  *Tamaño:* 6.838.715 bytes | *MD5:* C61E591EDD8BCAB9D789949D471FF617 | *SHA-256:* FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6 | *Atributo:* Solo Lectura
- **Archivo ACCDB:** H:\ResidenciaBD\Plantilla\GoldenMaster\Residencia_BE_GOLDEN_MASTER.accdb
  *Tamaño:* 6.733.824 bytes | *MD5:* A5B86926058CCF73DEDA7F0C8F83A323 | *SHA-256:* DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89 | *Atributo:* Solo Lectura
