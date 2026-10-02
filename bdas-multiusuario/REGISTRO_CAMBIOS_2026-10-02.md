# Registro de Modificaciones y Tareas — 2 de Octubre de 2026

## Archivos de Respaldo Generados en la Sesión
- `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-02_0016.xlsm`
- `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-02_0016.accdb`
- `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-02_0100.xlsm`
- `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-02_0112.accdb`
- `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-02_0113.accdb`
- `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-02_2234.xlsm`

---

## 1. Hitos Alcanzados en la Sesión

### A. Diagnóstico Forense del Plan de Optimización de Excel
- **Evaluación de la Acción 1 (16.383 columnas)**: Se constató mediante análisis de código XML interno que `RESIDENCIA GIJÓN` y `RESIDENCIA SOTO` **no poseen celdas más allá de la columna 30 (`AD`)**. Las columnas 31 a 16.384 están configuradas intencionadamente como `hidden="1"` para delimitar visualmente el libro. El peso de ~8 MB en estas hojas se debe a las fórmulas vivas en `U`, `AB` y `AC` extendidas hasta la fila 4.000 para futuras altas. Se determinó no alterar estas columnas ocultas para preservar la estética visual.
- **Evaluación de la Acción 2 (`calcChain.xml`)**: Se confirmó que el archivo ya no contiene la cadena de cálculo pesada anterior.
- **Ejecución de la Acción 3 (Rotación y purga de hojas LOG)**:
  - En la hoja `LOG_OVIEDO` se suprimieron los 6.637 registros antiguos de auditoría acumulados, preservando estrictamente la fila 1 de encabezados (`Usuario`, `Fecha/Hora`, `Nº ORDEN`, `COLUMNA`, `Valor anterior`, `Valor nuevo`).
  - La hoja pasó de **1,75 MB a 514 KB** de XML.
  - El libro `.xlsm` se redujo de **7,88 MB a 7,61 MB** (~270 KB comprimidos ahorrados de forma inmediata).
  - La función `ModuloLOG.RegistrarCambioLOG` reanudará la escritura limpia en la **fila 2**.

### B. Auditoría y Volcado Preventivo RGPD en Access (Acción 5)
- **Auditoría cruzada registro a registro**: Se contrastaron las hojas ocultas `BDAS GIJÓN`, `BDAS SOTO` y `BDAS OVIEDO` contra la tabla `Ordenes` en Access (`Residencia_BE.accdb`):
  - **GIJÓN**: 285 registros coincidentes (0 faltantes).
  - **SOTO**: 33 registros coincidentes (0 faltantes).
  - **OVIEDO**: 1.800 registros coincidentes (0 faltantes).
  - **Total**: 2.118 órdenes coincidentes al 100%.
- **Creación de tabla de respaldo preventivo**:
  - Para blindar la información histórica sin interferir con las búsquedas unificadas de `Ordenes_Historico`, se creó en Access la tabla **`BDAS_Historico_Snapshot`** que almacena las 2.118 órdenes completas con sus 41 columnas.
  - Las hojas `BDAS` en Excel se mantienen 100% intactas para no romper dependencias con `ModuloArchivadoDatosBDAS.bas`.

### C. Optimización de Fórmulas y Purga Masiva en Hojas RESUMEN (Acción 6 Completada)
- **Acotación de rangos VLOOKUP a `$A$1:$AE$4000`**:
  - Se sustituyeron las fórmulas que evaluaban la columna completa `$A:$AE` (1.048.576 filas) por el rango acotado `$A$1:$AE$4000` en `RESUMEN GIJÓN`, `RESUMEN SOTO` y `RESUMEN OVIEDO`.
  - Se subsanó un error latente en `RESUMEN OVIEDO` donde la fórmula de habitación buscaba erróneamente solo hasta la fila 1.303 (`$A$2:$AE$1303`), impidiendo ver habitaciones de órdenes posteriores.
- **Purga física de más de 90.000 fórmulas zombis**:
  - Se eliminaron las filas vacías que arrastraban fórmulas de búsqueda innecesarias hasta la fila ~3.980 en las tres hojas:
    - `RESUMEN GIJÓN`: reducida de 3.977 filas a **287 filas reales**.
    - `RESUMEN SOTO`: reducida de 3.981 filas a **34 filas reales**.
    - `RESUMEN OVIEDO`: reducida de 3.966 filas a **1.802 filas reales**.
- **Refactorización de módulos de actualización automática**:
  - Se actualizaron `ModuloResumenGijon.bas`, `ModuloResumenSoto.bas` y `ModuloResumenOviedo.bas` para medir la última fila real considerando tanto columna A como columna B, y limpiar automáticamente celdas huérfanas más allá de `contadorValidas + 1`.
- **Impacto y Ahorro**:
  - **XML Descomprimido (Carga de Memoria)**: reducido de **21,80 MB a 4,43 MB** (**-17,37 MB / -80% de memoria**).
  - **Tamaño del libro `.xlsm` en disco**: reducido de **7,61 MB a 6,17 MB** (**~1,44 MB de ahorro directo**).
  - Comportamiento de `Ctrl + Fin`: ahora se posiciona exactamente en la última orden real activa en todas las hojas RESUMEN.

---

### D. Descarte Definitivo de Acciones Menores (Acción 1 y Acción 4)
- **Acción 1 (16.383 columnas)**: Descartada definitivamente por análisis XML (no hay celdas vivas, son solo columnas ocultas para delimitar la estética de la hoja).
- **Acción 4 (Estilos en `styles.xml`)**: Descartada definitivamente tras análisis coste-beneficio. Con 1.375 estilos de celda (límite técnico en 64.000) y un ahorro marginal de ~300 KB, el riesgo de desfase en índices de formato XML y alertas de reparación en Excel no compensa la intervención.

---

## 2. Acción Estructural Pendiente: Acción 5 (Desacoplamiento de Hojas BDAS hacia Access)

### Pasos para Completar la Acción 5 (Fases 2 y 3):
1. **Refactorizar `ModuloArchivadoDatosBDAS.bas`**:
   - Sustituir la lógica de copia celda a celda hacia las hojas `BDAS GIJÓN`, `BDAS SOTO`, `BDAS OVIEDO` por sentencias SQL de archivado directo sobre Access (`Ordenes_Historico` o tabla histórica RGPD).
2. **Refactorizar `BorradoFilasHojasBDAS.bas` y `frmPanelRGPD.frm`**:
   - Conectar la purga de registros > 3 años para que ejecute `DELETE` sobre Access en lugar de operar sobre las hojas de Excel.
3. **Retirar las hojas `BDAS` de Excel**:
   - Una vez desacoplado y validado el código, eliminar `BDAS GIJÓN`, `BDAS SOTO` y `BDAS OVIEDO` del libro Excel.
   - **Ahorro esperado**: **~2,05 MB de XML descomprimido** (~400-500 KB comprimidos adicionales en disco, llevando el archivo a ~5,7 MB).

---

## 3. Hoja de Ruta y Orden de Prioridades Recomendado

Para garantizar la estabilidad operativa del sistema y evitar contaminación de diagnósticos, se establece la siguiente secuencia de trabajo:

```
[FASE 1: Inmediata - Hotfixes y Correcciones Funcionales]
  ├── Corregir adjudicación de números de factura (AsignarNumFactura, Facturación mensual)
  ├── Corregir registro y persistencia de auditoría LOG (ModuloLOG / LogActividad)
  └── Validar que todos los flujos de la operativa diaria funcionan al 100% sin bugs

[FASE 2: Estructural - Desacoplamiento de BDAS a Access (Acción 5)]
  ├── Refactorizar ModuloArchivadoDatosBDAS a sentencias SQL sobre Access
  ├── Conectar purga RGPD > 3 años con Access
  └── Eliminar las 3 hojas ocultas BDAS de Excel (peso final ~5,7 MB)
```

### Justificación Técnica y Operativa:
1. **Evitar la "Contaminación de Diagnóstico"**: Si se realiza primero el desacoplamiento estructural de BDAS y posteriormente se detecta un error de facturación o logs, resultará difícil discernir si el bug era previo o fue introducido por el refactor.
2. **Prioridad de Negocio / Integridad Contable**: Un fallo en número de factura o logs afecta directamente a la actividad del día a día (clientes, cobros, auditoría). El archivado de reservas > 90 días no tiene urgencia inmediata sobre las reservas activas.
3. **Suelo Firme para la Acción 5**: Al estar la facturación y los logs perfectamente corregidos y estables, la refactorización de `ModuloArchivadoDatosBDAS` se ejecutará sobre una base limpia y 100% testeable.


