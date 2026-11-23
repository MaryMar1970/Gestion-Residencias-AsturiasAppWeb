# Mapa de dependencias arquitectónicas — v1.0 rev1

## Alcance y cautela

Este documento resume las dependencias y contratos de la arquitectura objetivo definidos en `ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md`. **No es un mapa de procedimientos actualmente verificados en el XLSM de `main`.** Para una implementación, contrastar cada contrato con el código y los llamadores de la versión exacta identificada por hash.

El mapa as-built anterior del 10/11 se conserva en `docs/historico/MAPA_DEPENDENCIAS_2026-11-10.md`. Los informes 1H y 1Q–1X son evidencias fechadas y deben leerse con sus rectificaciones.

## 1. Identidad y ciclo de vida del expediente

- Identidad permanente: `IdInstanciaOrden`, GUID de 128 bits en formato canónico de 36 caracteres.
- No confundirla con `Ordenes.Id` (autonumérico local Access) ni con `IdOperacion` (idempotencia de una operación).
- La ubicación arquitectónica elegida para la identidad en Excel es AJ, oculta. La decisión final de v1.0 prevalece sobre ubicaciones candidatas examinadas en informes anteriores.
- A inicia el borrador; B (Fecha de Petición) determina `AnioOrden = Year(FechaPeticion)`.
- Access confirma transaccionalmente la reserva de `(Residencia, AnioOrden, NumOrden)`, protegida por índice único.
- No tratar un borrador como expediente confirmado ni actualizar Access como tal antes de la confirmación.

## 2. Persistencia, sincronización y concurrencia

- El expediente se dirige por `IdInstanciaOrden` en altas, actualizaciones, borrados e histórico.
- El borrador debe sobrevivir a una caída mediante WAL local durable y recuperable por GUID.
- Los reintentos son idempotentes: no crean otra instancia ni duplican operaciones.
- La sincronización offline y la recuperación siguen el contrato SICO-N2/WAL de la especificación.
- La identidad de sesión/instancia, el lease y los controles de concurrencia no sustituyen a la identidad permanente del expediente.

## 3. Facturación

- `AnioFactura = Year(FechaSalida)`; no usar `Year(Now)`, `FechaPeticion` ni el reloj local.
- Una orden corresponde a una factura física; no hay facturas agrupadas.
- `RegistroNumerosFactura` es el registro permanente del consumo.
- Clave única de consumo: `(Residencia, AnioFactura, NumFactura)`.
- `IdOperacion` asegura idempotencia y es distinto de `IdInstanciaOrden`.
- Un número consumido no se libera al borrar o archivar la asociación vigente.

## 4. Histórico y borrado

- La clave del histórico es `IdInstanciaOrden`, no `(Residencia, AnioOrden, NumOrden)`.
- El archivo combina upsert del histórico y eliminación del activo por GUID dentro de una única transacción Access.
- Se aplica el umbral de 17 campos cumplimentados y los nueve campos obligatorios especificados.
- Las diferencias entre sedes y la matriz C/P deben resolverse conforme a la especificación, no mediante inferencias de rangos.

## 5. Campos funcionales adicionales

- No realizar backfill: no existía persistencia previa que deba reconstruirse.
- `REPARTIR COSTE`: `dbText(2)`, valores `NO`/`SI`.
- `TAG` y `R`: `dbInteger`, valores enteros estrictos 1–9999; cero, negativos, decimales y texto inválidos.
- La adenda de Preflight es la corrección vinculante del dominio.

## 6. Evidencia as-built y límites

- La Orden 1H identifica la procedencia del XLSM 09/11 (SHA `B911…`) y `modDatabase.bas` (SHA `EF4F…`).
- Los informes 1Q–1X documentan inspecciones de código y dependencias; no certifican compilación ni comportamiento en producción.
- La inspección de Preflight se hizo sobre el backend de producción en modo de solo lectura, según su adenda. Las otras estaciones no fueron inspeccionadas.
- La preparación Gate 2 documenta un laboratorio aislado; no es autorización de despliegue.
- El código de `main` sigue asociado al XLSM del 22/10. La conciliación/promoción del candidato 09/11 es un paso separado y requiere autorización.

## Regla para futuros cambios

Antes de tocar un procedimiento, construir un submapa as-built con su firma, llamadores, dependencias, tablas/campos, hojas/rangos y pruebas. Si falta una fuente, solicitarla y no concluir por conjetura.