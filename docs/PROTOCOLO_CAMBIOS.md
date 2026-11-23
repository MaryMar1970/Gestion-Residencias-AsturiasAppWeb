# PROTOCOLO DE CAMBIOS — ResidenciaApp

## 1. Principio rector

La arquitectura objetivo de la aplicación Excel/VBA + Access se define en `docs/ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md`. Las rectificaciones específicas del Preflight se encuentran en `docs/RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md`. Este protocolo regula cómo inspeccionar, implementar, probar y promover cambios; no sustituye la especificación.

Distinguir siempre:

- **Diseño objetivo:** lo que exige la especificación.
- **As-built:** lo que el código y el esquema reales contienen en una versión identificada por hash.
- **Verificado:** lo que se ha ejecutado y probado con evidencia reproducible.

No presentar una decisión de diseño como si estuviera implantada o probada.

## 2. Antes de empezar

1. Leer `docs/INDICE_DOCUMENTAL.md`, la especificación, este protocolo y las evidencias pertinentes.
2. Confirmar rama, `HEAD`, remoto, `git status`, diferencias y stash existentes.
3. Identificar el XLSM y backend exactos mediante tamaño y SHA-256. No inferir versión por fecha física del archivo.
4. Confirmar el alcance: análisis, propuesta, laboratorio o producción. Una petición de análisis no autoriza editar ni probar el libro.
5. Comprobar procesos Excel/Access antes de operar sobre copias binarias.
6. Crear respaldo verificable y rama de trabajo cuando se autorice implementación. No alterar Golden Master ni `BASELINE`.

## 3. Inspección de dependencias

Antes de proponer un cambio, verificar en el código real:

- procedimiento, función, evento y todos sus llamadores;
- variables globales, módulos, formularios y controles;
- hojas, CodeNames, tablas, columnas, rangos, nombres definidos y fórmulas;
- conexiones DAO/ADO, consultas, transacciones, índices y restricciones;
- rutas de sincronización, WAL, bloqueos, recuperación e idempotencia.

No inferir dependencias por el nombre de un módulo. Si falta una fuente necesaria, detener la conclusión y solicitarla.

## 4. Diseño y ejecución

- ChatGPT actúa como arquitecto: define solución, contratos, algoritmo, errores, secuencia y aceptación.
- Antigravity actúa como ejecutor: inspecciona el repositorio real, implementa solo lo autorizado y entrega diffs y pruebas.
- Agrupar cambios de tamaño medio por dependencias; cada paquete debe ser autosuficiente y tener alcance cerrado.
- Verificar el texto exacto que se va a sustituir. Si no coincide, no forzar el parche: informar y detener esa sustitución.
- No mezclar cambios de laboratorio con producción ni promover código automáticamente.

## 5. Reglas VBA y entorno

- Compatibilidad: Excel 2016 en español, ACE/DAO de 32 bits.
- Módulos VBA en Windows-1252 y CRLF, sin BOM UTF-8 ni caracteres corruptos.
- No usar `Chr(211)` ni construcciones equivalentes para encubrir una codificación incorrecta.
- No generar texto español mediante VBScript.
- Mantener el flujo de trabajo y la disposición de hojas/formularios salvo autorización explícita.
- Optimizar operaciones masivas; evitar bucles celda a celda, esperas indefinidas y bloqueos de interfaz.

## 6. Validación

1. Ejecutar pruebas dirigidas en laboratorio, con datos de prueba y evidencia de entrada/salida.
2. Verificar concurrencia, duplicados, reintentos, pérdida de red, recuperación WAL y consistencia cuando el cambio afecte a esos contratos.
3. Compilar el proyecto en el entorno objetivo Excel 2016/ACE 32-bit y exigir cero errores antes de considerar válida una entrega.
4. Verificar no regresión de las sedes Gijón, Soto y Oviedo y de las funciones afectadas.
5. Registrar pruebas no ejecutadas, límites y estaciones no inspeccionadas. No convertir una inspección documental en certificación.

## 7. Git, trazabilidad y entrega

- Antes de editar: copia de seguridad, rama y baseline identificada.
- Después de editar: `git diff --check`, revisión completa del diff, hashes, inventario de archivos y estado Git.
- Registrar fecha, alcance, archivos, decisiones, pruebas, resultados y reversión.
- Commits descriptivos; sincronizar con `origin/main` solo después de la revisión y aceptación.
- No esconder cambios en stash ni declarar el árbol depurado por el mero hecho de que `git status` esté vacío. Cada archivo debe tener una decisión de conservar, archivar o eliminar.

## 8. Puerta de producción

La producción es `H:\ResidenciaBD\Residencia_BE.accdb` y el libro operativo identificado en la línea base. No ejecutar DDL, DML, migraciones, sustituciones de XLSM ni importaciones VBA en producción sin autorización expresa y una puerta de despliegue aprobada. La preparación de Gate 2 o un Preflight de solo lectura no equivalen a autorización de producción. La Fase 2A del Ledger no está aprobada. El parche WAL y el código de laboratorio no se aplican a producción.

## 9. Contratos que requieren especial protección

- Identidad del expediente por `IdInstanciaOrden` GUID; no confundir con `Ordenes.Id` ni `IdOperacion`.
- `AnioOrden = Year(FechaPeticion)`; `AnioFactura = Year(FechaSalida)`.
- Una orden → una factura física; consumo de número permanente en `RegistroNumerosFactura`.
- Histórico y borrado por GUID, con transacción única y umbrales definidos en la especificación.
- WAL local durable, sincronización idempotente y concurrencia conforme a v1.0 rev1.
- `TAG` y `R`: enteros estrictos 1–9999; no backfill de los cinco campos funcionales.