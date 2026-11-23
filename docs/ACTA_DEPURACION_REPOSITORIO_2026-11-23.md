# Acta de depuración del repositorio — 23/11/2026

## 1. Alcance

Saneamiento documental y de estructura del repositorio `H:\ResidenciaApp`. No se autorizó ni se realizó modificación del libro Excel, módulos VBA, backend Access, esquema ni datos de producción. No se ejecutó compilación VBA ni prueba funcional.

## 2. Punto de partida y salvaguardas

- Rama `main`, commit previo `61615b5`, sincronizado con `origin/main`.
- El stash revisado fue `stash@{0}` (`PRE_LIMPIEZA_Y_AUDITORIA_20261123_20261123`), basado en `9f67b2c`.
- Respaldo externo preservado: `H:\ResidenciaApp_RESPALDO_PRE_SYNC_20261123_201324`.
- Rama de seguridad preservada: `respaldo/main-antes-sincronizacion_20261123_201324`.
- Etiqueta Golden Master preservada: `GOLDEN_MASTER_2026-10-24`.
- Ninguno de estos elementos de recuperación se elimina como parte de esta depuración.

## 3. Especificación arquitectónica canónica

Se incorpora como documento rector único:

`docs/ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md`

- Tamaño: 51.176 bytes.
- SHA-256: `32639E57499B3EB47602F50508AB79E7FA2098FA0DF4C185588F0C10DB42925F`.
- La copia del stash y la copia del respaldo externo coinciden byte a byte. Las dos copias localizadas en la Biblioteca también tenían texto idéntico.
- La copia duplicada `(1)` no se incorpora.

La adenda de Preflight queda como corrección vinculante de los puntos que identifica; los informes de inspección se conservan como evidencia, no como autoridad arquitectónica alternativa.

## 4. Verificación de preservación del stash

### Archivos rastreados

- Diferencia del stash frente a su base: 170 rutas.
- 151 archivos modificados comprobados contra el respaldo externo: **151 coincidentes, cero ausentes y cero discrepancias** mediante hash de objeto Git.
- 19 eliminaciones del árbol del stash identificadas. Las eliminaciones ya presentes en `61615b5` se mantienen; las restantes se han revisado por categoría. No se ha eliminado código VBA ni binarios de producción.

### Árbol no rastreado

- Entradas del tercer árbol del stash: 1.466.
- 1.461 coincidencias verificadas en el respaldo externo.
- Cuatro diarios WAL sintéticos de Gate 2 no aparecen en el respaldo externo y deben recuperarse selectivamente desde el stash.
- El único aviso adicional de hash afectaba a `PARCHE_WAL_AISLADO.diff`; la comparación SHA-256 directa de ambos archivos confirmó que son idénticos. Fue un falso positivo de normalización, no una pérdida.

### Laboratorio Gate 2

Los dos archivos de `BASELINE` coinciden exactamente entre el laboratorio físico y el respaldo externo:

| Archivo | Tamaño | SHA-256 |
|---|---:|---|
| `BASELINE/BDAS_Multiusuario.xlsm` | 6.621.963 bytes | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` |
| `BASELINE/Residencia_BE.accdb` | 7.639.040 bytes | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` |

`BASELINE` no se modifica. El respaldo externo contiene 104 archivos en `EVIDENCIAS` y 45 en `TRABAJO`; el árbol del stash contiene 92 y 49, respectivamente. La unión prevista es 2 de `BASELINE` + 104 de `EVIDENCIAS` + 49 de `TRABAJO` = **155 archivos**.

Los cuatro diarios pendientes de copiar a `TRABAJO/WAL_TEST/` son:

| Ruta relativa bajo `TRABAJO/` | Bytes | SHA-256 |
|---|---:|---|
| `WAL_TEST/ESC_BORRADORES/SYS_COLA_SYNC_SALÓN-PC.wal` | 412 | `5E6754414E584E409F1F46FC157A29CF127C30FB2A5AC5B2FE0DFA4B305433BA` |
| `WAL_TEST/ESC_CIERRE_REAPERTURA/SYS_COLA_SYNC_SALÓN-PC.wal` | 255 | `1F47CB24447F2181D46448804BE4AE646209C265005ABC1143A93C06706DA470` |
| `WAL_TEST/ESC_FIFO_TEST/SYS_COLA_SYNC_SALÓN-PC.wal` | 465 | `775D085B0A492B29F3EBE668179ED45AAA9436CC319CFA84252E05007A826CD8` |
| `WAL_TEST/SYS_COLA_SYNC_SALÓN-PC.wal` | 74 | `A8B1C97A98A10966331CE1CFE4CE437903D62775EB8551C4389A2097A2DD4FC4` |

**Estado de recuperación:** pendiente de que el usuario ejecute la operación elevada solicitada para reconstruir `EVIDENCIAS` y `TRABAJO`. Después se copiarán estos cuatro WAL, se comprobarán los 155 archivos y solo entonces se retirará el stash.

## 5. Decisiones documentales

### Se conserva en la documentación activa

- Especificación arquitectónica v1.0 rev1.
- Informe Preflight v1.0 y su adenda.
- Rectificación 1H, informes 1Q–1X y complemento 1Q-A.
- Informes de preparación Gate 2 y adenda de seguridad/WAL.
- Registro de cambios 20/11 e índice de referencia 10/11.
- Registros e índices cronológicos anteriores, como trazabilidad.
- Auditoría `AUDITORIA_PANEL_CONTROL_ARCHIVADO_RGPD.md`.

### Se archiva como antecedente

- Documento maestro v0.2.7 y diseño técnico v0.2, superados por v1.0 rev1.
- Línea base y mapa de dependencias del 10/11; diccionario de rangos del 17/10.
- Inventarios JSON y `customUI.xml` como snapshots históricos, no fuentes activas.
- Informe de código 02/11 y PDF de cambios 04/08.

### Se elimina o no se reintroduce en el árbol activo

- Todos los `.txt` que quedaban en el repositorio; verificación física final: cero `.txt` fuera de dependencias y carpetas generadas excluidas.
- Borradores AI de Opus de agosto, el DOCX de recuperación automática y accesos directos `.txt` obsoletos.
- Paquete Fase 1, superado por el Paquete Técnico Maestro.
- Copia `GEMINI 4 DATOS`, que contenía el mismo ZIP anónimo que `transferencia_tecnica`.
- Carpetas de staging extraídas, exportaciones duplicadas, metadatos sueltos, scripts históricos de generación/auditoría de paquetes y clon Access de pruebas antiguo.
- Carpeta/ZIP de diagnóstico de lentitud como paquete binario dentro de Git; el informe de auditoría se conserva en `docs/evidencias/`.

Los elementos retirados del árbol activo siguen preservados en el respaldo externo. Los scripts de transferencia inspeccionados eran generadores/auditores de paquetes con rutas históricas fijas; algunos podían sobrescribir o eliminar bases de datos de demostración. No son dependencias de ejecución de ResidenciaApp.

## 6. Divergencia de versión que queda abierta

- XLSM de `main`: 6.755.115 bytes, SHA-256 `9114EA53C4DBA8EC2D88166B27F356C67CD75468990B61471801C6DB2B5626B9`, snapshot 22/10.
- Candidato 09/11 del respaldo: 6.621.963 bytes, SHA-256 `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`.
- `modDatabase.bas` del candidato: 298.794 bytes, SHA-256 `EF4FDB01AE65E46129057903BEF0DF8125A398CAC39A305B4E92438043C3BA47`, coincidente con la evidencia 1H.

No se ha sustituido el libro ni promovido ningún módulo. El candidato 09/11 y los cambios de código del stash se preservan fuera del árbol activo y no se declaran compilados ni probados. La reconciliación y eventual promoción requieren una decisión y autorización separadas.

## 7. Configuración y estructura

- Se restablecen reglas explícitas de `.gitattributes`, `.editorconfig` y `.vscode/settings.json`: VBA Windows-1252/CRLF; Markdown, JSON y XML UTF-8/LF; binarios protegidos.
- `README.md` raíz describe el repositorio y separa Excel/Access del prototipo web.
- `webapp/README.md` documenta el prototipo web por separado y elimina instrucciones obsoletas que apuntaban a un SQL inexistente.
- Se actualizan `AGENTS.md`, `PROTOCOLO_CAMBIOS.md`, línea base, mapa de dependencias y diccionario de rangos; se crea este índice documental.

## 8. Límites de certificación

No se modificó producción. Preflight fue una inspección documentada de solo lectura sobre el backend de producción; otras estaciones remotas no se inspeccionaron. Gate 2 Paso 1 no equivale a Gate 2 completo. La Fase 2A del Ledger no está aprobada y el parche WAL es exclusivamente de laboratorio. No se ejecutó compilación ni prueba funcional en esta depuración.

## 9. Cierre Git

- Commit documental inicial: `1d7fb17` (`chore: consolidate architecture docs and repository baseline`), publicado en `origin/main`.
- Verificación Git posterior: `HEAD = origin/main = 1d7fb17`; divergencia `0 0`. No hay cambios rastreados pendientes; `_AUDIT_STAGING/` permanece como único directorio local no rastreado porque contiene los cuatro WAL que aún deben recuperarse.
- Pendiente tras la restauración del laboratorio: copiar y verificar los cuatro WAL, comprobar los 155 archivos, eliminar `_AUDIT_STAGING/`, retirar `stash@{0}` y publicar el commit final de cierre. El respaldo externo, la rama de seguridad y Golden Master permanecen preservados.