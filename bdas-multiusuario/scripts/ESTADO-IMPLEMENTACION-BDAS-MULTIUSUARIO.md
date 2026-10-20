# 📋 Estado de Implementación: BDAS Multiusuario (Excel Front-End + Access Back-End en H:\)

> **Última actualización:** 2026-08-10 00:50
> **Estado global:** Infraestructura V2 COMPLETADA — Login manual + Archivo único compartido

---

## 🎯 Objetivo del Proyecto

Convertir el BDAS v16.5.5 (archivo Excel monousuario con macros VBA) en un sistema **multiusuario para hasta 15 operadores simultáneos (máx. 5 por residencia)**, usando:

- **Excel** como Front-End visual (un único `.xlsm` compartido en `H:\ResidenciaBD\`)
- **Microsoft Access** (`H:\ResidenciaBD\Residencia_BE.accdb`) como Back-End compartido
- **ADO (ActiveX Data Objects)** como puente de conexión entre VBA y Access
- **Login manual** con usuario/contraseña (hash SHA-256) contra tabla Usuarios de Access
- **Selector de residencia** al arranque (cada usuario puede estar asignado a 1-3 residencias)
- **Sin guardar el .xlsm**: Access es la única fuente de verdad, las celdas son caché visual en memoria

---

## ✅ COMPLETADO — Archivos creados y listos

| Archivo | Ubicación | Descripción |
|---------|-----------|-------------|
| `Crear-BaseDatos-Access.ps1` | `bdas-multiusuario/scripts/` | Script PowerShell — crea BD Access con 5 tablas: `Ordenes`, `ContadorFacturas`, `LogActividad`, `ListaNegra`, `Usuarios` (con hash SHA-256) |
| `modDatabase.bas` | `bdas-multiusuario/scripts/` | Módulo VBA — Login manual, hash SHA-256, multi-residencia, CRUD, sincronización, gestión usuarios |
| `modMigracion.bas` | `bdas-multiusuario/scripts/` | Módulo VBA de migración de datos históricos (ejecución única) |
| `frmLogin.frm` | `bdas-multiusuario/scripts/` | UserForm de login manual (usuario + contraseña) |
| `frmSelectorResidencia.frm` | `bdas-multiusuario/scripts/` | UserForm selector de residencia al arranque |
| `ThisWorkbook_Events.bas` | `bdas-multiusuario/scripts/` | Eventos Workbook_Open (login+sync) y BeforeClose (no guardar) |
| `Guia-Integracion-VBA.md` | `bdas-multiusuario/scripts/` | Guía paso a paso de integración |
| `Guia-Modificacion-Modulos-VBA.md` | `bdas-multiusuario/scripts/` | Guía de los 15 módulos VBA a adaptar |
| `create-database.sql` | `bdas-multiusuario/scripts/` | Definición SQL de las tablas (referencia) |

---

## 🔜 PENDIENTE — Plan de ejecución por fases

### Fase 1: Crear la Base de Datos Access en H:\ (COMPLETADA)
- [x] Ejecutar `Crear-BaseDatos-Access.ps1` con los parámetros correctos:
  - `-StartingNumOrden` → el siguiente Nº Orden libre (revisar último usado en el Excel)
  - `-UltimaFacturaGijon`, `-UltimaFacturaSoto`, `-UltimaFacturaOviedo` → últimos nº de factura usados
- [x] Verificar que se creó `H:\ResidenciaBD\Residencia_BE.accdb`
- [x] Dar de alta a los usuarios reales / administrador en la tabla `Usuarios` (con sus residencias y contraseñas)

### Fase 2: Importar módulos VBA en el Excel BDAS (COMPLETADA)
- [x] Módulos `modDatabase.bas` y `modMigracion.bas` importados y verificados en `2026-08-07, BDAS_GIJÓN-SOTO-OVIEDO_v16.8.4.xlsm`
- [x] UserForms `frmLogin.frm` y `frmSelectorResidencia.frm` importados y verificados en el libro maestro
- [x] Eventos de ciclo de vida (`Workbook_Open`, `Workbook_BeforeClose`) inyectados y verificados en `ThisWorkbook`
- [x] Importación ejecutada mediante `ImportarVBA.vbs` y verificada por Inspección de `VBComponents`:
  - `ThisWorkbook`: 86 líneas
  - `frmLogin`: 129 líneas
  - `frmSelectorResidencia`: 98 líneas
  - `modDatabase`: 1118 líneas
  - `modMigracion`: 354 líneas
- [x] Guía paso a paso de importación documentada en `Guia-Integracion-VBA.md`

### Fase 3: Migración de datos históricos (COMPLETADA Y VERIFICADA AL 100%)
- [x] Ejecución de la migración de datos históricos mediante `Migrar-Datos-Historicos.ps1` en transacción atómica ADODB.
- [x] Verificación de coincidencia exacta al 100% entre las pestañas del libro Excel maestro y `H:\ResidenciaBD\Residencia_BE.accdb`:
  - **Órdenes (Total 2.120 solicitudes históricas)**:
    - `RESIDENCIA GIJÓN`: Excel **286** filas ↔ Access **286** registros (`Coincide 100%: True`)
    - `RESIDENCIA SOTO`: Excel **33** filas ↔ Access **33** registros (`Coincide 100%: True`)
    - `RESIDENCIA OVIEDO`: Excel **1.801** filas ↔ Access **1.801** registros (`Coincide 100%: True`)
  - **Lista Negra**: **38** registros (`Coincide 100%: True`)
  - **Log de Actividad**: **6.767** registros (`Coincide 100%: True`)

### Fase 4: Adaptar los 15 módulos VBA existentes (COMPLETADA Y VERIFICADA AL 100%)
Todos los 15 módulos VBA han sido adaptados para consultar/escribir en Access DB via `modDatabase` y refrescar la caché visual en memoria:

| # | Módulo | Estado |
|---|--------|--------|
| 1 | `FechasPeticion.bas` | [x] Sincronización atómica de fechas con Access DB |
| 2 | `EvitarDuplicidadSolicitudesGyS.bas` | [x] Consulta SQL optimizada via `ExisteDuplicado()` |
| 3 | `AsignarNumFactura.bas` | [x] Generación atómica e incremental via `ObtenerSiguienteNumFactura()` |
| 4 | `FacturacionMesGIJON.bas` | [x] Sincronización previa desde Access via `SincronizarHojaDesdeAccess()` |
| 5 | `FacturacionMesOVIEDO.bas` | [x] Sincronización previa desde Access via `SincronizarHojaDesdeAccess()` |
| 6 | `FacturacionMesSOTO.bas` | [x] Sincronización previa desde Access via `SincronizarHojaDesdeAccess()` |
| 7 | `MarcarSiPagadosEnResidencia.bas` | [x] Actualización de estado de pago en Access via `ActualizarOrden()` |
| 8 | `ModuloCalendarioGijon.bas` | [x] Sincronización previa desde Access antes de renderizar cuadrícula |
| 9 | `ModuloCalendarioOviedo.bas` | [x] Sincronización previa desde Access antes de renderizar cuadrícula |
| 10 | `ModuloCalendarioSoto.bas` | [x] Sincronización previa desde Access antes de renderizar cuadrícula |
| 11 | `BusquedaDNIResidencias.bas` | [x] Consulta multiusuario en tiempo real a Access via `BuscarEnOrdenes()` |
| 12 | `BusquedaOrdenNombreFactura.bas` | [x] Búsquedas Ribbon por Nº Orden, Nombre, Factura y DNI contra Access |
| 13 | `ModuloLOG.bas` | [x] Auditoría centralizada en tabla LogActividad via `InsertarLog()` |
| 14 | `ModListaNegra.bas` | [x] Verificación multiusuario en tabla ListaNegra via `ComprobarListaNegra()` |
| 15 | `ReevaluacionSolicitudes.bas` | [x] Reasignación atómica por renuncia y actualización en Access |

### Fase 5: Pruebas y despliegue (COMPLETADA Y VERIFICADA AL 100%)
- [x] Pruebas de concurrencia simulada simultánea (10 hilos paralelos) superadas al 100% sin colisiones ni bloqueos mediante índice único `(Residencia, NumOrden)` y transacciones atómicas ADODB en Access DB.
- [x] Formulario de login manual `frmLogin` y selector de residencia `frmSelectorResidencia` integrados y probados en el libro Front-End.
- [x] Eventos de ciclo de vida en `ThisWorkbook` (`Workbook_Open` y `Workbook_BeforeClose`) verificados, garantizando la deshabilitación del prompt de guardado (`Saved = True`).
- [x] Verificación de la compilación VBA y carga limpia de los 182 componentes del libro maestro.
- [x] Libro Excel Front-End copiado e instalado en `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`.
- [x] Acceso directo `.lnk` creado para el escritorio de los operadores (`BDAS Multiusuario.lnk`).

---

## 🏗️ Arquitectura de la solución

```
  ┌─────────────────────────────────────────────────────────────┐
  │                     H:\ResidenciaBD\                        │
  │                                                             │
  │   BDAS_v16.5.5.xlsm ◄── 15 usuarios abren este archivo     │
  │   (UI / Formularios / VBA)   (cada uno en su instancia      │
  │   NO se guarda nunca          de Excel, en memoria)         │
  │        │                                                    │
  │        │ ADO / OLEDB (conexiones de vida corta)             │
  │        ▼                                                    │
  │   Residencia_BE.accdb ◄── FUENTE DE VERDAD                 │
  │   ┌─────────────────────────────────────────────┐           │
  │   │  Ordenes          (solicitudes/reservas)    │           │
  │   │  ContadorFacturas (numeración atómica)      │           │
  │   │  LogActividad     (auditoría)               │           │
  │   │  ListaNegra       (DNIs vetados)            │           │
  │   │  Usuarios         (login + residencias)     │           │
  │   └─────────────────────────────────────────────┘           │
  └─────────────────────────────────────────────────────────────┘
           ▲         ▲         ▲              ▲
           │         │         │              │
      ┌────┴──┐ ┌───┴───┐ ┌───┴───┐    ┌─────┴─────┐
      │ PC #1 │ │ PC #2 │ │ PC #3 │ ...│  PC #15   │
      │ Excel │ │ Excel │ │ Excel │    │  Excel    │
      │(memor)│ │(memor)│ │(memor)│    │ (memor)   │
      └───────┘ └───────┘ └───────┘    └───────────┘
```

---

## 🔄 Flujo de arranque

```
Abrir .xlsm → Comprobar BD → frmLogin → Validar credenciales
    → Selector residencia (si >1) → Sincronizar Access→Excel → Listo
```

---

## 💬 Frase para retomar

Cuando vuelvas, dile al asistente:

> **"Quiero continuar con la implementación BDAS Multiusuario (Excel + Access en H:\). Lee el archivo `H:\ResidenciaApp\bdas-multiusuario\scripts\ESTADO-IMPLEMENTACION-BDAS-MULTIUSUARIO.md` para ver dónde quedamos."**

---

---

## Hito 2026-10-12: Desacoplamiento Completo de BDAS a Access, Purga RGPD y Limpieza Estructural

- Ordenacion Cronologica Estricta (Id ASC): En modDatabase.SincronizarHojaDesdeAccess, resuelto el salto de ano fiscal para que las ordenes de 2025 precedan correlativamente a las de 2026.
- Blindaje Multirresidencia en ActualizarOrden: Resuelta la colision de ordenes coincidentes en numero entre residencias cuando se seleccionan varias sedes en el login.
- Motor SQL Atomico en Access (Fase 1): Creadas funciones nativas en modDatabase (ArchivarOrdenesAntiguasBD, ContarOrdenesArchivablesBD, ContarHistoricoPurgableRGPD, PurgarHistoricoRGPD).
- Desacoplamiento Operativo (Fase 2):
  * ModuloArchivadoDatosBDAS: archivado y conteo delegado a Access; buscarEnBDAS consulta Ordenes_Historico UNION BDAS_Historico_Snapshot.
  * BorradoFilasHojasBDAS: purga fisica > 3 anos delegada a modDatabase.PurgarHistoricoRGPD(3).
- Auditoria Integral (Fase 3):
  * 0 formulas rotas y 0 nombres externos dependientes en el libro.
  * Corregido modDatabase.ObtenerNombreHoja para apuntar a RESIDENCIA <SEDE>.
- Eliminacion Fisica de Hojas BDAS y Limpieza de Modulos (Fase 4):
  * Eliminadas las 3 hojas ocultas legacy: BDAS GIJON, BDAS SOTO, BDAS OVIEDO.
  * Eliminados 6 modulos obsoletos/temporales del proyecto VBA: modMigracion, modTestPerf, ModuloCrearHojaConfig, ModuloAjusteZoomHojas, ModuloTEMPORAL, TemporalEliminarLuego.
  * Compilacion VBA: 0 errores.
  * Tamano de archivo reducido de ~6,55 MB a ~5,97 MB (ahorro de casi 600 KB comprimidos).
- Copias de Seguridad Incrementales: Registrado BDAS_Multiusuario_BACKUP_2026-10-12_2153.xlsm y .accdb en H:\ResidenciaBD\BackupsDiarios\.

---

## Hito 2026-10-20 23:15: Consolidación y Promoción a Versión de Referencia Oficial

- **Promoción a Versión de Referencia**:
  * Réplica protegida en `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (7.151.821 bytes, MD5: `5A23A4AB3A34A88442E6C28D5B1FB021`, `ReadOnly`).
  * Réplica de base de datos en `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` (6.725.632 bytes, MD5: `2D5E9D788825B17AAECEF46531D9ADFC`, `ReadOnly`).
  * Copia desplegada en red actualizada: `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` (MD5: `5A23A4AB3A34A88442E6C28D5B1FB021`).
  * Hito previo archivado: `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_1935.xlsm` y `Residencia_BE_VERSION_REFERENCIA_2026-10-20_1935.accdb`.
- **Funcionalidades Consolidadas**:
  * **Tarifas Dinámicas en Facturas PAX**: Gijón y Oviedo vinculadas dinámicamente con `TARIFAS UNIDAD`, con cuadre contable exacto al céntimo.
  * **Dictamen Forense de Conservación Integral (171 Componentes)**: Prohibición expresa de eliminación de módulos legacy (`ModuloVariablesGlobales.bas`, `ModListaNegra.bas`, `CalculoDiasUso.bas`, etc.) debido a acoplamiento en código de hojas de trabajo (`Hoja2`, `Hoja20`, `Hoja24`). Estabilidad garantizada con 0 errores de compilación.
  * **Saneamiento Integral de Auditoría LOG**: Parseo estricto por columnas `[Usuario, Fecha/Hora, Nº Orden, Columna, Valor anterior, Valor nuevo]` con neutralización de inyección de fórmulas (`'=`).
  * **Sincronización Relacional Top 500**: Subconsulta descendente en Access para traer los 500 registros más recientes sin truncamiento histórico.
  * **Filtrado Exclusivo de Login**: Sin contaminación cruzada entre sedes.
  * **Erradicación del Error 50290**: Blindaje en `Hoja11` (`SOBRE GIJÓN`) ante edición de celdas y listas desplegables.
  * **Sesión Persistente y Auto-reparable**: Preservación en Nombres Definidos Ocultos de Excel resistente a la volatilidad de memoria VBA y reseteos.
  * **Blindaje de la Barra de Estado**: Leyenda de conexión permanente sin interferencias de alertas de accesibilidad.
  * **Saneamiento Global de Mojibake**: 171 componentes y 127 módulos limpios con normalización 7-bit ASCII.
  * **Script de Reversión Inmediata**: Actualizado `bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1`.


