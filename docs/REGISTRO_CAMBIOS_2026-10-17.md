# Registro de Modificaciones y Tareas — 17 de Octubre de 2026

## Hitos y Cambios Realizados en la Sesión

### A. Creación del Sistema Integral de Documentación y Gobernanza Técnica
Se ha diseñado y generado el marco completo de control de cambios, dependencias y arquitectura técnica en el directorio `H:\ResidenciaApp\docs\`:

1. **[MAPA_DEPENDENCIAS.md](file:///H:/ResidenciaApp/docs/MAPA_DEPENDENCIAS.md):**
   - Mapeo de los 69 callbacks de la cinta personalizada RibbonX (`customUI.xml`).
   - Mapeo de eventos globales de `ThisWorkbook` (`Workbook_Open`, `Workbook_SheetDeactivate`) y eventos de hoja (`Worksheet_Change` en `Hoja34_TarifasUnidad.cls` y hojas de residencia).
   - Grafo de llamadas hacia el núcleo de persistencia `modDatabase.bas` y seguridad `ModuloConfigSegura.bas`.
   - Matriz de interacción CRUD con las 8 tablas de la base de datos Access `Residencia_BE.accdb`.

2. **[DICCIONARIO_RANGOS.md](file:///H:/ResidenciaApp/docs/DICCIONARIO_RANGOS.md):**
   - Catálogo de las 49 hojas del libro Excel con su `Name`, `CodeName` inmutable de VBA y visibilidad.
   - Detalle columna por columna de `RESIDENCIA GIJÓN` (30 columnas, importe en `U`), `RESIDENCIA SOTO` (30 columnas, importe en `U`) y `RESIDENCIA OVIEDO` (33 columnas, importe en `V` debido a la columna adicional previa de cama supletoria).
   - Estructura y fórmulas maestras matriciales con `SUMAPRODUCTO` para el cálculo dinámico de tarifas e IVA.
   - Registro de los 90 rangos nombrados (`Defined Names`) en Excel.

3. **[LINEA_BASE_REFERENCIA.md](file:///H:/ResidenciaApp/docs/LINEA_BASE_REFERENCIA.md):**
   - Certificado de integridad con hashes criptográficos MD5 verificados:
     - `BDAS_Multiusuario.xlsm`: `9E9685D7F5F8E33B8A139EE4929AE543`
     - `Residencia_BE.accdb`: `740739C2D02DD8D98F17C4DD75F09AD2`
   - Inventario físico desagregado de los 141 archivos VBA en `vba-modules`.
   - Esquema relacional de las 8 tablas de la base de datos Access.
   - Puntos de restauración disponibles (Versión de Referencia y Golden Master).

4. **[PROTOCOLO_CAMBIOS.md](file:///H:/ResidenciaApp/docs/PROTOCOLO_CAMBIOS.md):**
   - Flujo obligatorio de 6 pasos para cualquier cambio:
     1. Análisis de Impacto Previo.
     2. Copia de Seguridad Previa Obligatoria (`PRE`).
     3. Codificación con Reglas de Oro (`GIJÓN` literal, borrado exclusivo RGPD, no contraseñas hardcodeadas, `UserInterfaceOnly:=True`).
     4. Compilación y Validación Sintáctica Limpia en VBA.
     5. Copia de Seguridad Posterior (`POST`) y comprobación de Hash MD5.
     6. Registro y Actualización Documental.
   - Procedimiento estandarizado de rollback y contingencia.

### B. Adopción de la Nueva Versión de Referencia y Archivo Activo
Por instrucción expresa del usuario, se fija como:
- **Versión de Referencia Oficial y Archivo Activo de Trabajo:**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`
  - Hash MD5 inicial: `57FA43E5E28820D5853F9D0154BDB2AE`
  - Timestamp de modificación: `17/10/2026 00:39:18`
- Todos los análisis, propuestas de modificación e implementaciones se realizarán directa y exclusivamente sobre este archivo.

### C. Corrección y Blindaje Definitivo de Fórmulas en Sincronización Multirresidencia (GIJÓN, SOTO, OVIEDO)
- **Incidencia detectada por el usuario:**
  - Al iniciar sesión en `RESIDENCIA SOTO`, desaparecían las fórmulas de las columnas `Q`, `R` y `AB` desde la fila 2 hasta la 35 (a partir de la 36 permanecían intactas).
  - En `RESIDENCIA OVIEDO`, desaparecían las fórmulas de la columna `AF` desde la fila 2 hasta la 1801.
- **Diagnóstico técnico y causa raíz:**
  - En `modDatabase.bas` (`SincronizarHojaDesdeAccess`), la macro limpiaba por bloques rangos que incluían columnas de fórmulas (`A:T` y `V:AB` en Soto/Gijón, y columna 32 `AF` en Oviedo).
  - La macro volcaba datos de Access asumiendo que `Q` y `R` eran números de habitaciones en Soto (cuando son fórmulas de precios y suplementos) y que `AB`/`AF` (`GRABACIÓN SOLICITUD`) eran marcas temporales de auditoría `FechaGrabacion` (cuando en Excel almacenan la fórmula generadora del nombre de archivo PDF).
  - Las filas más allá de las órdenes existentes en Access (34 órdenes en Soto, 1800 en Oviedo) conservaban la fórmula al no ser procesadas por el bucle de sincronización.
- **Solución implementada:**
  1. **Limpieza selectiva (`ClearContents`):** Exclusión estricta de las 11 columnas de fórmulas de las 3 residencias.
  2. **Volcado selectivo de datos:** Exclusión de sobreescritura de `Q`, `R` en Soto, y de `AB`/`AF` en Gijón, Soto y Oviedo.
  3. **Blindaje y auto-reparación universal post-sincronización:** Comprobación al vuelo fila a fila en las 11 columnas de fórmula; si alguna celda no contiene fórmula, se localiza la fórmula maestra en la fila plantilla (mediante escaneo inteligente hacia adelante) y se reinyecta en formato `FormulaR1C1`.
- **Validación empírica:**
  - Compilación limpia del proyecto VBA (0 errores).
  - Ejecución de simulación de sincronización completa en las 3 residencias: 100% de celdas con `HasFormula = Verdadero` y cálculo exacto.
- **Respaldos de sesión asociados:**
  - `PRE`: `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-17_2312_PRE_BLINDAJE_FORMULAS.xlsm` (MD5: `57FA43E5E28820D5853F9D0154BDB2AE`)
  - `POST`: `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-17_2316_POST_BLINDAJE_FORMULAS.xlsm` (MD5: `F814E071502409B94A89C2FA545AC513`)
  - Archivo de red sincronizado y verificado: `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` (MD5: `F814E071502409B94A89C2FA545AC513`)

---

## Próximo Paso Inmediato
- Arrancar la **Fase 1: Centralización de Acciones Privilegiadas por Rol** siguiendo estrictamente el nuevo protocolo:
  1. Generar backup previo `BDAS_Multiusuario_BACKUP_2026-10-17_XXXX_PRE_FASE1_ROLES.xlsm`.
  2. Implementar `ValidarPermisoAdmin(nombreAccion)` en módulo central.
  3. Desmantelar los `InputBox` manuales de contraseña (Categoría A).
