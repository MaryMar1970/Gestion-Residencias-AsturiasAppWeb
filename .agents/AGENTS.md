# AGENTS.md — ResidenciaApp

## Alcance del repositorio

- `bdas-multiusuario/` contiene la aplicación operativa Excel/VBA y sus exportaciones.
- El backend Access de producción está fuera del repositorio: `H:\ResidenciaBD\Residencia_BE.accdb`.
- `webapp/` es un prototipo separado .NET/React. No es la autoridad funcional de la aplicación Excel/Access.

## Lectura obligatoria antes de intervenir

1. `docs/INDICE_DOCUMENTAL.md`.
2. `docs/ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md`.
3. `docs/RESIDENCIAAPP_ADENDA_PREFLIGHT_v1_0.md` y `docs/RESIDENCIAAPP_INFORME_PREFLIGHT_v1_0.md` cuando afecte a datos, esquema o puertas.
4. `docs/PROTOCOLO_CAMBIOS.md`.
5. Evidencias 1H y 1Q–1X que afecten al flujo concreto.

## Autoridad y precedencia documental

- La especificación arquitectónica v1.0 rev1 es la referencia principal para la arquitectura objetivo, decisiones y contratos de Excel/Access.
- La adenda de Preflight es vinculante para sus rectificaciones concretas: dominio de `TAG`/`R` 1–9999, identificación del backend consultado y límites de la inspección.
- Los informes de auditoría son evidencias de inspecciones fechadas; no prueban por sí solos que el código actual implemente el diseño.
- `RESIDENCIAAPP_DOCUMENTO_MAESTRO_v0.2.7`, el diseño v0.2, los mapas y diccionarios históricos son antecedentes, no autoridades superiores a v1.0 rev1.
- Separar siempre arquitectura objetivo, código realmente inspeccionado y comportamiento probado. No inferir dependencias por nombres.

## Invariantes arquitectónicos que no deben quebrantarse

- `IdInstanciaOrden` es la identidad permanente del expediente (GUID); no confundirlo con `Ordenes.Id` autonumérico ni con `IdOperacion`.
- La identidad técnica se dirige por GUID; la ubicación arquitectónica elegida es AJ, oculta. Verificar el libro y los llamadores reales antes de implantarla.
- Alta A → B: A inicia borrador; B determina `AnioOrden = Year(FechaPeticion)`; Access confirma la reserva transaccional de `(Residencia, AnioOrden, NumOrden)`. No actualizar Access como expediente confirmado antes de completar la confirmación.
- Facturación: `AnioFactura = Year(FechaSalida)`. Un expediente corresponde a una factura física; no hay facturas agrupadas. `RegistroNumerosFactura` mantiene el consumo único `(Residencia, AnioFactura, NumFactura)`; un número consumido no se libera.
- Histórico y borrado se dirigen por `IdInstanciaOrden`; el archivo debe ser transaccional y respetar el umbral de 17 campos y los nueve obligatorios definidos en la especificación.
- Los cinco campos funcionales no admiten backfill. `REPARTIR COSTE`: `dbText(2)`, `NO`/`SI`. `TAG` y `R`: `dbInteger`, enteros estrictos 1–9999; cero, negativos, decimales y texto son inválidos.
- La sincronización offline requiere WAL local durable, idempotencia y los contratos de concurrencia/identidad definidos en v1.0 rev1. No sustituirlos por heurísticas locales.

## Seguridad y límites de autorización

- No modificar el libro Excel, módulos VBA, backend Access, esquema ni datos de producción sin autorización expresa para ese cambio concreto.
- No ejecutar DDL/DML ni repetir consultas de Preflight sobre producción. La adenda aclara que la inspección documentada apuntó al backend de producción en modo de solo lectura; otras estaciones remotas no fueron inspeccionadas.
- No confundir el cierre del Paso 1 de preparación Gate 2 con la aprobación completa de Gate 2 ni con autorización de despliegue.
- La Fase 2A del Ledger no está aprobada. El parche WAL y el código modificado en Gate 2 son exclusivamente de laboratorio: no aplicarlos a producción.
- La carpeta `LABORATORIO_GATE2_20261120_1558/` es local y excluida de Git. `BASELINE/` es inmutable; nunca alterar sus ACL ni sobrescribir sus archivos.
- Preservar la etiqueta `GOLDEN_MASTER_2026-10-24`, la rama `respaldo/main-antes-sincronizacion_20261123_201324` y el respaldo externo `H:\ResidenciaApp_RESPALDO_PRE_SYNC_20261123_201324`.

## Entorno y reglas de código

- Compatibilidad obligatoria: Excel 2016 en español, ACE/DAO de 32 bits.
- VBA: Windows-1252 nativo y CRLF; no introducir UTF-8/BOM en módulos VBA ni mojibake. No usar sustituciones como `Chr(211)` para ocultar errores de codificación.
- No generar texto español mediante VBScript. Mantener comentarios y mensajes en español correctamente codificados.
- Verificar nombres reales de procedimientos, eventos, formularios, controles, hojas, tablas, columnas y rangos antes de cambiar código.
- No introducir formularios de progreso ni cambios de interfaz ajenos al alcance autorizado.

## Método de trabajo

- En análisis solamente: inspeccionar y dictaminar; no editar ni probar el libro.
- En implementación: ChatGPT define arquitectura, algoritmo, contratos y criterios de aceptación; Antigravity ejecuta únicamente el paquete autorizado; el usuario importa, compila y prueba en Excel 2016 por etapas.
- Agrupar cambios por dependencias en paquetes técnicos autosuficientes. Hacer copia previa, trabajar en laboratorio, aplicar cambios quirúrgicos y verificar el contexto exacto antes de sustituir código.
- Si falta un módulo, llamador o dependencia necesaria, solicitarlo; no inventar ni completar por conjetura.
- Exigir `git diff`, hashes, inventario de archivos, evidencia de pruebas y registro cronológico. Fail-fast ante discrepancias.
- Nunca usar un stash como sustituto de una depuración real. Antes de retirar un stash, clasificar y preservar o descartar conscientemente cada contenido.
- No declarar compilación, pruebas o validaciones que no se hayan ejecutado.

## Discrepancia de versión que sigue abierta

El XLSM de `main` es el snapshot del 22/10/2026 (SHA-256 `9114EA53C4DBA8EC2D88166B27F356C67CD75468990B61471801C6DB2B5626B9`). El candidato del 09/11/2026 (SHA-256 `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7`) y sus exportaciones se preservan en el respaldo externo. No promover ni sustituirlos automáticamente: requieren una decisión y autorización separadas.