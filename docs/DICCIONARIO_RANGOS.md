# Diccionario de rangos — estado de verificación

No existe en este repositorio, a fecha de esta depuración, un diccionario completo de rangos certificado contra el XLSM más reciente.

- El último inventario detallado conservado es `docs/historico/DICCIONARIO_RANGOS_2026-10-17.md`; es una fotografía histórica, no una fuente vigente.
- Los inventarios `sheet_details.json` y `architecture_summary.json` también se han archivado como snapshots históricos.
- La especificación v1.0 rev1 reserva AJ para `IdInstanciaOrden`; esta es una decisión de arquitectura objetivo, no una prueba de que el libro versionado en `main` ya lo implemente.
- Los informes 1H–1X aportan evidencias parciales de rangos y dependencias. La ubicación candidata AO examinada en un informe anterior no sustituye la decisión final AJ de la especificación.

Antes de modificar una celda, rango, nombre definido, fórmula, filtro, ordenación o `ScrollArea`, inspeccionar el libro exacto, su CodeName, el evento que lo modifica y todos los procedimientos que lo leen o escriben. No copiar rangos ni asumir equivalencias entre Gijón, Soto y Oviedo por similitud visual.