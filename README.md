# ResidenciaApp — repositorio principal

Este repositorio reúne dos líneas de trabajo distintas. No deben confundirse ni tratarse como una única implementación.

## 1. Aplicación operativa Excel / VBA / Access

- Libro: `bdas-multiusuario/BDAS_Multiusuario.xlsm`.
- Backend de producción: `H:\ResidenciaBD\Residencia_BE.accdb` (externo al repositorio).
- Entorno de referencia: Excel 2016 en español, ACE/DAO de 32 bits.
- **Documento rector:** [`docs/ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md`](docs/ResidenciaApp_Especificacion_Arquitectonica_v1_0_rev1.md).
- **Jerarquía, estado y evidencias:** [`docs/INDICE_DOCUMENTAL.md`](docs/INDICE_DOCUMENTAL.md).

La especificación define la arquitectura objetivo y sus contratos. No significa que todas sus decisiones estén ya implantadas en el código del libro.

## 2. Prototipo web

`webapp/` contiene un prototipo .NET/React separado. Sus instrucciones están en [`webapp/README.md`](webapp/README.md). No se debe asumir que este prototipo es la aplicación operativa ni que está conectado al backend Access de producción.

## Estado y precauciones

- Antes de cualquier trabajo, revisar [`docs/LINEA_BASE_REFERENCIA.md`](docs/LINEA_BASE_REFERENCIA.md) y ejecutar `git status`.
- El XLSM versionado en `main` corresponde al snapshot del 22/10/2026. Existe un candidato posterior del 09/11/2026 preservado en el respaldo externo; no se ha promovido durante la depuración porque esta tarea no autoriza modificar el libro ni los módulos VBA.
- `LABORATORIO_GATE2_20261120_1558/` es local y está excluido de Git. `BASELINE/` es inmutable.
- No modificar el backend de producción, el libro Excel ni módulos VBA sin autorización específica y sin superar las puertas de validación correspondientes.