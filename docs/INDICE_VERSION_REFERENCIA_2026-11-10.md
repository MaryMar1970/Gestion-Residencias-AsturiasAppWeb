# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: Hito 2026-11-10 15:45 (Blindaje de Celdas Combinadas en Turnos-Precios (B25:D25), Optimización de Rendimiento de Carga de Tarifas (ScreenUpdating/Calculation), Persistencia Multiusuario de Tarifas y Precios en Access (ConfigTarifasCeldas con 291 registros relacionales y segundo tramo 2026) y Cero Mojibake)

---

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`  
  *Tamaño:* 6,621,963 bytes | *MD5:* `DF455E2A5D9F215DD7E97DA8E0E335FC` | *SHA-256:* `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`

- **Copia Timestamped en VersionReferencia (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-10_1545.xlsm`  
  *Tamaño:* 6,621,963 bytes | *MD5:* `DF455E2A5D9F215DD7E97DA8E0E335FC` | *SHA-256:* `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`

- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`  
  *Tamaño:* 7,614,464 bytes | *MD5:* `D9BB14ABB808DE0255E47F7A1AD3CC2C` | *SHA-256:* `EBCD61EBB64416A009F35F29C2A2B6F03E2CD10890CBAD2A55546949E214DCFD`

- **Copia Timestamped Base de Datos (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA_2026-11-10_1545.accdb`  
  *Tamaño:* 7,614,464 bytes | *MD5:* `D9BB14ABB808DE0255E47F7A1AD3CC2C` | *SHA-256:* `EBCD61EBB64416A009F35F29C2A2B6F03E2CD10890CBAD2A55546949E214DCFD`

- **Copia Activa de Trabajo (Desarrollo y Pruebas):**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6,621,963 bytes | *MD5:* `DF455E2A5D9F215DD7E97DA8E0E335FC` | *SHA-256:* `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`

- **Copia Desplegada en Red (Producción Multiusuario Operativa):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`  
  *Tamaño:* 6,621,963 bytes | *MD5:* `DF455E2A5D9F215DD7E97DA8E0E335FC` | *SHA-256:* `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`

---

### 2. Matriz de Trazabilidad: Funcionalidades vs. Copias Incrementales de Seguridad

A continuación se detalla la correspondencia estricta entre las intervenciones técnicas realizadas y las copias de seguridad incrementales generadas en `H:\ResidenciaBD\BackupsDiarios\`:

| Backup Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Finalidad Técnica y Alcance Operativo |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.xlsm` | 6.590.683 B | `CD9A878DB02B874BF17A4F2F9EDA3BA6` | `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F` | **Línea Base Previa de Partida.** Incorpora subsistema consolidado de persistencia de bloqueos ACID. |
| `Residencia_BE_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.accdb` | 6.754.304 B | `5816C92401274DD2A84EB238AB92980E` | `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40` | Respaldo previo de base de datos Access de partida. |
| `Residencia_BE_BACKUP_2026-11-09_1939_PRE_ADAPT_TARIFAS.accdb` | 6.832.128 B | `8AE298A62024E5B1919864E3D52FFEB0` | `826C6543D698DC894458EB185D95C246995642504938D4DF142EE301138F97E4` | **Punto de Control Pre-Adaptación BD.** Estado previo al reajuste de celdas en Access. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_1939_PRE_ADAPT_TARIFAS.xlsm` | 5.863.175 B | `ACDB945D813470FEBF7C873B00CD6FA0` | `BE2329E25828D890BF74CE4F3E417A295CE167B67EBBC89DB92D3E3F58C0843F` | Respaldo previo del libro Excel antes de la adaptación. |
| `Residencia_BE_BACKUP_2026-11-09_1939_POST_ADAPT_TARIFAS.accdb` | 7.507.968 B | `F058EF2C9D8E3462EEC180836719ABD7` | `52804682882208D26428567260CDBD1DEF287F55FA4CA5401909EF1C04753EED` | **Punto de Control Post-Adaptación BD.** 155 registros relacionales en `TARIFAS UNIDAD` y 291 totales. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_1939_POST_ADAPT_TARIFAS.xlsm` | 5.863.175 B | `ACDB945D813470FEBF7C873B00CD6FA0` | `BE2329E25828D890BF74CE4F3E417A295CE167B67EBBC89DB92D3E3F58C0843F` | Respaldo post-adaptación de base de datos. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_2000_VERSION_REFERENCIA.xlsm` | 6.598.187 B | `EB89934732A313E09A2028020438933E` | `2C5E7D888D16C9F13B9DF9C24B1FD126254E206AD52D4355F56221E64649F6B7` | **Hito Intermedio Referencia 20:00.** Persistencia de tarifas y cobertura de rangos `E3:N9` y `E13:N19`. |
| `Residencia_BE_BACKUP_2026-11-09_2000_VERSION_REFERENCIA.accdb` | 7.507.968 B | `F61C12C9E65E1BADE30B90E561787D9A` | `5811CE37B5D309635C2B4C570FF1285C7ECD7463A1BCE3BAE24B5F985B431E95` | Base de datos intermedia Hito 20:00 con 291 registros de tarifas. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_2050_VERSION_REFERENCIA.xlsm` | 6.621.963 B | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | Cierre previo con blindaje contra celdas combinadas `B25:D25`. |
| `Residencia_BE_BACKUP_2026-11-09_2050_VERSION_REFERENCIA.accdb` | 7.540.736 B | `CBCA7AF72BE3DF9243530CD487D5E53B` | `340721144FF888488A82013CAE837D6EB3DBE718FBE4645F2095D238AAB69F37` | Base de datos Hito 20:50. |
| `BDAS_Multiusuario_BACKUP_2026-11-10_1545_VERSION_REFERENCIA.xlsm` | 6,621,963 B | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **Versión de Referencia Oficial Definitiva Consolidada.** Libro maestro íntegro con 173 componentes VBE, 0 errores de compilación, cero mojibake y 100% verificado contra producción. |
| `Residencia_BE_BACKUP_2026-11-10_1545_VERSION_REFERENCIA.accdb` | 7,614,464 B | `D9BB14ABB808DE0255E47F7A1AD3CC2C` | `EBCD61EBB64416A009F35F29C2A2B6F03E2CD10890CBAD2A55546949E214DCFD` | **Base de Datos Referencia Oficial Consolidada.** Tablas operativas completas, `ConfigTarifasCeldas` con 291 registros e índice único verificado, incluyendo tramo tarifario activo 2026. |

- **Histórico de Referencias Previas en `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`:**
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-09_2050.xlsm` (6.621.963 B | MD5: `DF455E2A5D9F215DD7E97DA8E0E335FC`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-11-09_2050.accdb` (7.540.736 B | MD5: `CBCA7AF72BE3DF9243530CD487D5E53B`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-09_2000.xlsm` (6.598.187 B | MD5: `EB89934732A313E09A2028020438933E`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-11-09_2000.accdb` (7.507.968 B | MD5: `F61C12C9E65E1BADE30B90E561787D9A`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-06_0045.xlsm` (6.590.683 B | MD5: `CD9A878DB02B874BF17A4F2F9EDA3BA6`)
  - `Residencia_BE_VERSION_REFERENCIA_2026-11-06_0045.accdb` (6.754.304 B | MD5: `5816C92401274DD2A84EB238AB92980E`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-03_0015.xlsm` (6.616.947 B | MD5: `06E31C80A08574CA44F016850FBA6655`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_2202.xlsm` (6.385.789 B | MD5: `549921707952BC6EAA423A6CBC107500`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0330.xlsm` (6.430.806 B | MD5: `F73F5C628319611FD8C56CA773E60B67`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-02_0110.xlsm` (6.838.431 B | MD5: `33AC497F6CF682AFCB79DDAAA07446C2`)
  - `BDAS_Multiusuario_VERSION_REFERENCIA_2026-11-01_0425.xlsm` (6.809.197 B | MD5: `5727B66025C7278C95B5596CA22BD351`)

---

### 3. Inventario de Componentes y Mejoras Clave del Hito 2026-11-10 15:45

1. **`modDatabase.bas` (Función `CargarTarifasDesdeBD`):**
   - Blindaje de celdas combinadas: Verificación de `celTarget.MergeCells` y descarte seguro de celdas esclavas con `GoTo SiguienteCeldaCargar`.
   - Control de rendimiento: Bloqueo de redibujado (`Application.ScreenUpdating = False`) y suspensión de recálculos (`Application.Calculation = xlCalculationManual`) con restauración simétrica garantizada.
   - Definición canónica exacta de rangos en `ObtenerRangosEditablesTarifas`: `Array("A3:C7", "E3:N9", "E13:N19")` para `TARIFAS UNIDAD`, y `Array("A1:F20", "A22:F23", "B25:D31")` para `TURNOS-PRECIOS`.
   - Control optimista de concurrencia mediante `CalcularFirmaHojaTarifas` y `CalcularFirmaTarifasDesdeBD` con serialización inyectiva determinista (`R<fila>C<col>:<tipo>:<len>:<val>;`).
   - Transacciones atómicas unificadas en `GuardarTarifasEnBD` con detección de colisiones concurrentes y registro de auditoría en `LogActividad`.

2. **`Residencia_BE.accdb` (`ConfigTarifasCeldas`):**
   - Tabla relacional con 291 registros (155 correspondientes a `TARIFAS UNIDAD` y 136 a `TURNOS-PRECIOS`).
   - Índice único `idx_Tarifas_HojaCelda` sobre `(NombreHoja, Fila, Columna)` blindando contra duplicados.
   - Cobertura de tarifas 2026 con segundo tramo a partir del 01/06/2026.

3. **`ThisWorkbook.cls` (Manejo Interactivo de Colisiones):**
   - Hook en `Workbook_SheetDeactivate` que evalúa `TieneCambiosPendientesHojaTarifas(nomSh)`.
   - En caso de colisión de edición concurrente con otro operador, notifica interactivamente y permite la recarga inmediata de las tarifas del servidor sin pérdida de datos.

4. **`BotonTarifasUnidadDescanso.bas` (Sincronización Previa de Visualización):**
   - Métodos `MostrarTarifaUnidad` y `MostrarTarifaDescanso` que invocan `CargarTarifasDesdeBD` antes de visualizar la hoja, impidiendo desincronizaciones en puestos cliente.
