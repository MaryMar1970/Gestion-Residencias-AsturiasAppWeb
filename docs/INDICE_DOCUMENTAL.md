# Índice documental y jerarquía de autoridad

## 1. Núcleo arquitectónico

1. **[Especificación arquitectónica v1.0 rev1](ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md)** — documento rector de la arquitectura objetivo, identidad, ciclo A→B, persistencia, facturación, histórico, concurrencia, contratos y puertas.
2. **[Adenda de Preflight v1.0](RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md)** — rectificaciones vinculantes a los resultados de Preflight, en particular el dominio de `TAG`/`R` y la identificación del backend consultado.
3. **[Informe Preflight v1.0](RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md)** — evidencia base; debe leerse junto con su adenda, que corrige puntos expresamente identificados.

La especificación es la autoridad de diseño. No debe confundirse con la implementación actual ni con una autorización de despliegue.

## 2. Estado y protocolo

- [Línea base actual](LINEA_BASE_REFERENCIA.md) — identifica el commit, los hashes y la divergencia entre el XLSM de `main` y el candidato posterior 09/11.
- [Mapa de dependencias arquitectónicas](MAPA_DEPENDENCIAS.md) — resume contratos objetivo; exige reconstruir el mapa as-built antes de cada cambio.
- [Diccionario de rangos](DICCIONARIO_RANGOS.md) — explica el estado de verificación y remite al inventario histórico.
- [Protocolo de cambios](PROTOCOLO_CAMBIOS.md) — regula análisis, implementación, pruebas, Git y puertas de producción.
- [Acta de depuración 23/11](ACTA_DEPURACION_REPOSITORIO_2026-11-23.md) — registra las decisiones de conservación, archivo y eliminación y la verificación del stash/respaldo.
- [Registro de cambios 20/11](REGISTRO_CAMBIOS_2026-11-20.md) — registro cronológico más reciente disponible.
- [Índice de versión de referencia 10/11](INDICE_VERSION_REFERENCIA_2026-11-10.md) — fotografía histórica de esa referencia; no describe el XLSM actualmente versionado en `main`.

## 3. Evidencias técnicas que se conservan

- **Identidad y órdenes:** [rectificación 1H](RESIDENCIAAPP_INFORME_RECTIFICACION_ORDEN_1H.md), [1Q](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1Q.md), [complemento 1Q-A](RESIDENCIAAPP_INFORME_COMPLEMENTARIO_ORDEN_1Q_A.md), [1R](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1R.md), [1S](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1S.md), [1T](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1T.md), [1U](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1U.md), [1V](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1V.md), [1W](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1W.md) y [1X](RESIDENCIAAPP_INFORME_EVIDENCIAS_ORDEN_1X.md).
- **Preflight:** [informe base](RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md) y [adenda correctiva](RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md).
- **Gate 2:** [informe de preparación del laboratorio](INFORME_PASO1_SETUP_LABORATORIO.md) y [adenda de seguridad/WAL](INFORME_ADENDA_PASO1_SEGURIDAD_Y_WAL.md).
- **RGPD y rendimiento:** [auditoría del panel RGPD](evidencias/AUDITORIA_PANEL_CONTROL_ARCHIVADO_RGPD.md). El ZIP diagnóstico completo se conserva en el respaldo externo, no se incorpora a Git.
- **Trazabilidad:** [registro de cambios 20/11](REGISTRO_CAMBIOS_2026-11-20.md) e [índice de referencia 10/11](INDICE_VERSION_REFERENCIA_2026-11-10.md). Los demás índices y registros fechados son historia cronológica, no documentos de autoridad concurrentes.

## 4. Antecedentes archivados

En `historico/` se conservan, con valor histórico y sin autoridad sobre v1.0 rev1:

- [Documento maestro v0.2.7](historico/arquitectura/RESIDENCIAAPP_DOCUMENTO_MAESTRO_v0.2.7.md) y [diseño técnico v0.2](historico/arquitectura/DISENO_TECNICO_ARQUITECTURA_OBJETIVO_RESIDENCIAAPP_v0.2_2026-11-19.md): antecedentes de arquitectura, superados por v1.0 rev1.
- [Línea base 10/11](historico/LINEA_BASE_REFERENCIA_2026-11-10.md), [mapa de dependencias 10/11](historico/MAPA_DEPENDENCIAS_2026-11-10.md) y [diccionario de rangos 17/10](historico/DICCIONARIO_RANGOS_2026-10-17.md): fotografías anteriores, no vigentes.
- [Inventarios JSON](historico/inventarios/snapshots_historicos/architecture_summary.json), [detalle de hojas](historico/inventarios/snapshots_historicos/sheet_details.json) y [customUI.xml](historico/inventarios/snapshots_historicos/customUI.xml): snapshots históricos; no son fuentes activas ni deben importarse al XLSM de `main`.
- [Informe de código 02/11](historico/informes/informe_codigo_2026-11-02.md) y [PDF de cambios 04/08](historico/2026-08-04_Cambios_ResidenciasAPP.pdf): antecedentes, no autoridad actual.

## 5. Paquetes y material fuera del repositorio

Los paquetes de transferencia Fase 1, Maestro y Anónimo, sus carpetas de staging, exportaciones duplicadas, scripts generadores y el clon de Access de pruebas se mantienen en el respaldo externo `H:\ResidenciaApp_RESPALDO_PRE_SYNC_20261123_201324`. No son dependencias de ejecución del sistema ni fuentes actuales de arquitectura. El paquete Fase 1 fue superado por el Maestro; `GEMINI 4 DATOS` era una copia exacta del ZIP anónimo. No se reintroducen en el árbol activo.

## 6. Separación del prototipo web

`webapp/` es un subproyecto .NET/React separado. Su guía es `webapp/README.md`. No debe confundirse con el sistema operativo Excel/Access ni suponerse una migración ya aprobada.

## 7. Bloqueos y límites conocidos

- El XLSM de `main` es el snapshot 22/10 (SHA-256 `9114EA53…`). El candidato 09/11 (SHA-256 `B9117659…`) y su `modDatabase.bas` (SHA-256 `EF4FDB01…`) se conservan en el respaldo externo; no se promovieron por falta de autorización para modificar XLSM/VBA en esta depuración.
- No se ejecutó compilación VBA ni prueba funcional durante esta limpieza.
- La adenda aclara que Preflight consultó el backend de producción en solo lectura; otras estaciones remotas no se inspeccionaron.
- El Paso 1 de preparación Gate 2 no equivale a Gate 2 completo ni a autorización de producción. La Fase 2A del Ledger no está aprobada; el parche WAL de laboratorio no se aplica a producción.
- `LABORATORIO_GATE2_20261120_1558/` es local y excluido de Git; `BASELINE/` es inmutable. La recuperación de `EVIDENCIAS/` y `TRABAJO/` requiere completar la operación elevada indicada en la conversación y restaurar cuatro WAL sintéticos desde el stash antes de retirarlo.