# REGISTRO DE CAMBIOS Y FUNCIONALIDADES
## Fecha: 2026-11-09 | Hito Oficial: Versión de Referencia 2026-11-09 20:50
### Ámbito: Blindaje de Celdas Combinadas en Turnos-Precios (B25:D25), Optimización de Rendimiento de Carga de Tarifas (ScreenUpdating/Calculation) y Persistencia Multiusuario de Tarifas y Precios en Access (ConfigTarifasCeldas)

---

### 1. Resumen de la Intervención y Contexto Técnico

En la sesión de trabajo del **09 de Noviembre de 2026**, se completó con éxito la auditoría forense, implementación, corrección de rangos, resolución de errores en celdas combinadas, optimización masiva de rendimiento de carga y consolidación formal del subsistema de **Persistencia y Gestión Multiusuario de Tarifas y Precios en Microsoft Access (`Residencia_BE.accdb`)**, estructurado en los siguientes pilares operativos:
1. **Blindaje de Celdas Combinadas en `Turnos-Precios` (`B25:D25`):** Corrección en `CargarTarifasDesdeBD` en [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L7078) para validar `celTarget.MergeCells` y saltar celdas secundarias combinadas (`C25`, `D25`) que no son la celda superior izquierda (`MergeArea.Cells(1,1)`), resolviendo de forma definitiva la excepción `"Esta acción no se puede realizar en una celda combinada"`.
2. **Optimización de Rendimiento de Carga de Tarifas:** Incorporación de control determinista de entorno (`Application.ScreenUpdating = False` y `Application.Calculation = xlCalculationManual`) con preservación y restauración simétrica de sus estados originales en `CargarTarifasDesdeBD`. Se elimina el cuello de botella de redibujado celda a celda en las 155 celdas de `TARIFAS UNIDAD`, reduciendo el tiempo de despliegue de ~30 segundos a ~1 segundo.
3. **Esquema Relacional e Índices de Tarifas (`ConfigTarifasCeldas`):** Estructura relacional en Access con clave primaria autonumérica e índice único `idx_Tarifas_HojaCelda` compuesto por `(NombreHoja, Fila, Columna)`. Total de registros relacionales: 291 (155 en `TARIFAS UNIDAD` + 136 en `TURNOS-PRECIOS`).
4. **Corrección Integral de Cobertura en `TARIFAS UNIDAD`:** Adaptación y blindaje de los rangos completos para GIJÓN (`E3:N9`) y OVIEDO (`E13:N19`), integrando las columnas E a N (Desde, Hasta, Individual, Ind-Comisión, Doble, Doble-Comisión, Cama Supletoria, Supletoria-Comisión, Res. Estudiantes Clase A y Clase B), junto a IVA `A3:C7`.
5. **Persistencia de `TURNOS-PRECIOS`:** Cobertura de rangos `A1:F20`, `A22:F23` y `B25:D31`, con omisión estricta de celdas con fórmulas de cálculo.
6. **Control Optimista de Concurrencia y Sincronización:** Serialización canónica determinista de celdas (`CalcularFirmaHojaTarifas` y `CalcularFirmaTarifasDesdeBD`), detección de colisiones (*Lost Updates*), recarga previa al abrir hojas y guardado condicional al salir (`Workbook_SheetDeactivate`).

---

### 2. Trazabilidad Estricta: Funcionalidades vs. Copias de Seguridad Incrementales

| Copia Incremental | Tamaño | Hash MD5 | Hash SHA-256 | Hito / Funcionalidad Técnica Vinculada |
| :--- | :--- | :--- | :--- | :--- |
| `BDAS_Multiusuario_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.xlsm` | 6.590.683 B | `CD9A878DB02B874BF17A4F2F9EDA3BA6` | `0995BF36B8BAA12465CFEE6261BDE91132C855CB3829ED9A853566A5BA4B819F` | **Línea Base Previa de Partida.** Bloqueos y ocultación ACID consolidada (OP-FINAL-BLOQ-01..03). |
| `Residencia_BE_BACKUP_2026-11-06_0045_VERSION_REFERENCIA.accdb` | 6.754.304 B | `5816C92401274DD2A84EB238AB92980E` | `3F6482EEA7F319599CB89A92FDE0312D118AB106FD6161C76AF4308EE2FF4A40` | Respaldo Access previo al subsistema de persistencia de tarifas. |
| `Residencia_BE_BACKUP_2026-11-09_1939_PRE_ADAPT_TARIFAS.accdb` | 6.832.128 B | `8AE298A62024E5B1919864E3D52FFEB0` | `826C6543D698DC894458EB185D95C246995642504938D4DF142EE301138F97E4` | **Punto de Control Pre-Adaptación BD.** Estado de la base de datos previo al reajuste de rangos en `ConfigTarifasCeldas`. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_1939_PRE_ADAPT_TARIFAS.xlsm` | 5.863.175 B | `ACDB945D813470FEBF7C873B00CD6FA0` | `BE2329E25828D890BF74CE4F3E417A295CE167B67EBBC89DB92D3E3F58C0843F` | Respaldo previo del libro Excel antes de la adaptación. |
| `Residencia_BE_BACKUP_2026-11-09_1939_POST_ADAPT_TARIFAS.accdb` | 7.507.968 B | `F058EF2C9D8E3462EEC180836719ABD7` | `52804682882208D26428567260CDBD1DEF287F55FA4CA5401909EF1C04753EED` | **Punto de Control Post-Adaptación BD.** `ConfigTarifasCeldas` con 155 registros en `TARIFAS UNIDAD` y 291 totales. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_1939_POST_ADAPT_TARIFAS.xlsm` | 5.863.175 B | `ACDB945D813470FEBF7C873B00CD6FA0` | `BE2329E25828D890BF74CE4F3E417A295CE167B67EBBC89DB92D3E3F58C0843F` | Respaldo post-adaptación de base de datos. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_2000_VERSION_REFERENCIA.xlsm` | 6.598.187 B | `EB89934732A313E09A2028020438933E` | `2C5E7D888D16C9F13B9DF9C24B1FD126254E206AD52D4355F56221E64649F6B7` | **Hito Intermedio Referencia 20:00.** Persistencia de tarifas y cobertura de rangos `E3:N9` y `E13:N19`. |
| `Residencia_BE_BACKUP_2026-11-09_2000_VERSION_REFERENCIA.accdb` | 7.507.968 B | `F61C12C9E65E1BADE30B90E561787D9A` | `5811CE37B5D309635C2B4C570FF1285C7ECD7463A1BCE3BAE24B5F985B431E95` | Base de datos intermedia Hito 20:00 con 291 registros de tarifas. |
| `BDAS_Multiusuario_BACKUP_2026-11-09_2050_VERSION_REFERENCIA.xlsm` | 6,621,963 B | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **Versión de Referencia Oficial Definitiva Consolidada.** Cierre con blindaje contra celdas combinadas `B25:D25`, optimización ScreenUpdating/Calculation en carga de tarifas, 173 componentes VBE, 0 errores de compilación y cero mojibake. |
| `Residencia_BE_BACKUP_2026-11-09_2050_VERSION_REFERENCIA.accdb` | 7,540,736 B | `CBCA7AF72BE3DF9243530CD487D5E53B` | `340721144FF888488A82013CAE837D6EB3DBE718FBE4645F2095D238AAB69F37` | **Base de Datos Referencia Oficial Consolidada.** Tablas operativas íntegras, 291 registros en `ConfigTarifasCeldas`, índices únicos verificados. |

---

### 3. Detalle Técnico de Funcionalidades y Mejoras Implementadas

#### A. Blindaje contra Excepciones en Celdas Combinadas (`B25:D25` en `Turnos-Precios`)
- **Problema subsanado:** Al pulsar el botón Ribbon `Descanso`, `CargarTarifasDesdeBD` recorría el rango `B25:D31` e intentaba ejecutar `ClearContents` o `.Value =` sobre `C25` y `D25`, provocando el error `"Esta acción no se puede realizar en una celda combinada"`.
- **Implementación integrada:**
  ```vba
  Set celTarget = wsTar.Cells(f, c)
  If celTarget.MergeCells Then
      If celTarget.Address <> celTarget.MergeArea.Cells(1, 1).Address Then
          rs.MoveNext
          GoTo SiguienteCeldaCargar
      End If
  End If
  ```
  Garantiza que solo se interactúe con la celda representativa superior izquierda (`B25`), ignorando con seguridad las celdas esclavas combinadas.

#### B. Optimización de Rendimiento en Carga de Tarifas (`ScreenUpdating` y `Calculation`)
- **Problema subsanado:** La carga de `TARIFAS UNIDAD` demoraba ~30 segundos al redibujar la pantalla y recalcular fórmulas en cada una de las 155 escrituras de celda.
- **Implementación integrada:**
  - Desactivación preventiva al inicio del bucle:
    ```vba
    screenOrig = Application.ScreenUpdating
    Application.ScreenUpdating = False
    calcOrig = Application.Calculation
    Application.Calculation = xlCalculationManual
    ```
  - Restauración garantizada en bloque de salida y en bloque de control de errores `ErrCargarTar`:
    ```vba
    Application.Calculation = calcOrig
    Application.ScreenUpdating = screenOrig
    ```
  - Reducción del tiempo de respuesta: de ~30 segundos a < 1 segundo.

#### C. Persistencia Multiusuario y Esquema de Tarifas en Access
- **Tabla relacional:** `ConfigTarifasCeldas` en [`Residencia_BE.accdb`](file:///H:/ResidenciaBD/Residencia_BE.accdb).
- **Campos:** `Id` (AutoInc PK), `NombreHoja` (TEXT 40), `Fila` (LONG), `Columna` (LONG), `TipoDato` (TEXT 10), `ValorNumero` (DOUBLE), `ValorFecha` (DATETIME), `ValorTexto` (TEXT 255), `UsuarioModificacion` (TEXT 50), `UltimaModificacion` (DATETIME).
- **Índice único:** `idx_Tarifas_HojaCelda` sobre `(NombreHoja, Fila, Columna)`.

#### D. Corrección de Rangos en `modDatabase.bas` (`ObtenerRangosEditablesTarifas`)
- **Implementación integrada:**
  ```vba
  Case "TARIFAS UNIDAD"
      ObtenerRangosEditablesTarifas = Array("A3:C7", "E3:N9", "E13:N19")
  Case "TURNOS-PRECIOS"
      ObtenerRangosEditablesTarifas = Array("A1:F20", "A22:F23", "B25:D31")
  ```

#### E. Control Optimista de Concurrencia y Sincronización Automática
- **Procedimientos:** `CalcularFirmaHojaTarifas` y `CalcularFirmaTarifasDesdeBD`.
- **Detección de colisiones:** Evaluación en `Workbook_SheetDeactivate` de [`ThisWorkbook.cls`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls), impidiendo la sobreescritura accidental (*Lost Update*) y ofreciendo recargar de inmediato.
- **Botones Ribbon:** `MostrarTarifaUnidad` y `MostrarTarifaDescanso` en [`BotonTarifasUnidadDescanso.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonTarifasUnidadDescanso.bas) sincronizan desde Access antes de mostrar la hoja.

---

### 4. Certificación de Calidad y Puertas de Aceptación

- [x] **Compilación VBE:** `Application.VBE.CommandBars.FindControl(1, 578).Enabled == False` (0 errores de compilación).
- [x] **Total Componentes VBE:** Exactamente 173 componentes en el proyecto Excel.
- [x] **Codificación Nativa:** Windows-1252 ANSI estricta en todos los archivos exportados en `vba-modules/`.
- [x] **Anti-Mojibake:** 0 patrones anómalos detectados por el guardián automatizado `Validar-Codificacion-VBA.py --all` (181 archivos analizados).
- [x] **Regla Estricta de Nombres:** Preservación de 'GIJÓN' como literal estricto.
- [x] **Inmutabilidad de Golden Master:** La carpeta `H:\ResidenciaBD\Plantilla\GoldenMaster\` no ha sido modificada ni accedida.
