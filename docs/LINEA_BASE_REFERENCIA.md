# Línea base de referencia — estado de la depuración 23/11/2026

## Repositorio Git

- Rama: `main`.
- Commit previo a esta depuración: `61615b5`.
- `origin/main` estaba sincronizada con ese commit antes de los cambios documentales.
- El commit base del stash revisado era `9f67b2c`.
- Rama de seguridad preservada: `respaldo/main-antes-sincronizacion_20261123_201324`.
- Etiqueta Golden Master preservada: `GOLDEN_MASTER_2026-10-24`.
- Respaldo externo que conserva el estado previo: `H:\ResidenciaApp_RESPALDO_PRE_SYNC_20261123_201324`.

## Divergencia del libro Excel

| Artefacto | Tamaño | SHA-256 | Procedencia |
|---|---:|---|---|
| `bdas-multiusuario/BDAS_Multiusuario.xlsm` de `main` | 6.755.115 bytes | `9114EA53C4DBA8EC2D88166B27F356C67CD75468990B61471801C6DB2B5626B9` | Snapshot del hito 22/10/2026; blob Git `7607819b3c75a1a6030ba66cc89ba1421f59f8d1` |
| Candidato `BDAS_Multiusuario.xlsm` del respaldo externo | 6.621.963 bytes | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | Snapshot del 09/11/2026, coherente con el índice de referencia 10/11 y los informes 1H–1X |

El candidato del 09/11 es posterior al libro versionado en `main`. No se ha promovido ni sustituido durante esta depuración: la autorización excluía modificar el libro Excel y los módulos VBA. Se conserva en el respaldo externo para una reconciliación posterior explícita.

El `modDatabase.bas` del candidato, en CRLF, tiene 298.794 bytes y SHA-256 `EF4FDB01AE65E46129057903BEF0DF8125A398CAC39A305B4E92438043C3BA47`, coincidente con la evidencia de la Orden 1H. Esto identifica la procedencia del código; no acredita compilación ni pruebas en Excel.

## Access y Preflight

- Backend de producción: `H:\ResidenciaBD\Residencia_BE.accdb`; no se modificó.
- La adenda de Preflight aclara que la inspección documentada apuntó a ese backend de producción en modo de solo lectura, no a un backend de pruebas.
- No se repitieron consultas en producción. Otras estaciones remotas siguen sin inspeccionar.
- Los datos y conclusiones del Preflight son evidencia histórica fechada, no una lectura en vivo ni una aprobación de migración.

## Laboratorio Gate 2

- Ruta local excluida de Git: `LABORATORIO_GATE2_20261120_1558/`.
- `BASELINE/` contiene el libro y el backend de referencia y debe permanecer inmutable.

| Archivo de BASELINE | Tamaño | SHA-256 |
|---|---:|---|
| `BDAS_Multiusuario.xlsm` | 6.621.963 bytes | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` |
| `Residencia_BE.accdb` | 7.639.040 bytes | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` |

La copia física de `BASELINE` coincide por SHA-256 con la copia externa verificada. Estos hashes identifican snapshots de laboratorio; no son una afirmación sobre el estado actual del backend de producción.
- El respaldo externo contiene 104 archivos en `EVIDENCIAS/` y 45 en `TRABAJO/`; el árbol del stash contiene 92 y 49, respectivamente. Los cuatro WAL sintéticos son las cuatro entradas de `TRABAJO/` que faltan en el respaldo externo. La unión que se debe recuperar es `BASELINE` (2) + `EVIDENCIAS` (104) + `TRABAJO` (49) = 155 archivos.
- La restauración de `EVIDENCIAS/` y `TRABAJO/` desde el respaldo externo, más los cuatro WAL sintéticos conservados en el stash, queda pendiente de completar y verificar mediante la operación elevada solicitada.
- La preparación del Paso 1 no certifica Gate 2 completo ni autoriza producción. La Fase 2A del Ledger no está aprobada; el parche WAL es solo de laboratorio.

## Alcance de esta depuración

No se modificaron el XLSM, los módulos VBA, el backend Access ni los datos de producción. No se ejecutó compilación ni prueba funcional. Los cambios realizados son documentales, de organización del repositorio y de configuración de texto.