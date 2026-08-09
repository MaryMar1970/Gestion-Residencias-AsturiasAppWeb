# 📋 Estado de Implementación: BDAS Multiusuario (Excel Front-End + Access Back-End en H:\)

> **Última actualización:** 2026-08-10 00:26
> **Estado global:** Fase de infraestructura COMPLETADA — Pendiente ejecución por fases

---

## 🎯 Objetivo del Proyecto

Convertir el BDAS v16.5.5 (archivo Excel monousuario con macros VBA) en un sistema **multiusuario para 6 operadores simultáneos**, usando:

- **Excel** como Front-End visual (cada operador tiene su copia local en el escritorio)
- **Microsoft Access** (`H:\ResidenciaBD\Residencia_BE.accdb`) como Back-End compartido en la unidad de red `H:\`
- **ADO (ActiveX Data Objects)** como puente de conexión entre VBA y Access
- **Patrón de doble escritura:** los datos se graban en Access (fuente de verdad) Y en la hoja local (caché visual para macros de calendario/impresión)

---

## ✅ COMPLETADO — Archivos creados y listos

| Archivo | Ubicación | Descripción |
|---------|-----------|-------------|
| `Crear-BaseDatos-Access.ps1` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Script PowerShell que crea la BD Access con tablas `Ordenes`, `ContadorFacturas`, `LogActividad`, `ListaNegra` e índices |
| `modDatabase.bas` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Módulo VBA (708 líneas) — Capa completa de acceso a datos ADO: inserción atómica, sincronización Access→Excel, búsquedas, lista negra, log, facturación, conexiones de vida corta |
| `modMigracion.bas` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Módulo VBA de migración de datos históricos de hojas Excel a Access (ejecución única) |
| `Guia-Integracion-VBA.md` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Guía paso a paso: crear BD, importar módulo, adaptar macros, desplegar a 6 PCs |
| `Guia-Modificacion-Modulos-VBA.md` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Guía detallada de los 15 módulos VBA existentes a adaptar, con código antes/después |
| `create-database.sql` | `H:\ResidenciaApp\bdas-multiusuario\scripts\` | Definición SQL de las tablas (referencia) |

---

## 🔜 PENDIENTE — Plan de ejecución por fases

### Fase 1: Crear la Base de Datos Access en H:\
- [ ] Ejecutar `Crear-BaseDatos-Access.ps1` con los parámetros correctos:
  - `-StartingNumOrden` → el siguiente Nº Orden libre (revisar último usado en el Excel)
  - `-UltimaFacturaGijon`, `-UltimaFacturaSoto`, `-UltimaFacturaOviedo` → últimos nº de factura usados
- [ ] Verificar que se creó `H:\ResidenciaBD\Residencia_BE.accdb`

### Fase 2: Importar módulos VBA en el Excel BDAS
- [ ] Abrir el Excel maestro → ALT+F11 → Importar `modDatabase.bas`
- [ ] Importar `modMigracion.bas`

### Fase 3: Migración de datos históricos
- [ ] Ejecutar `modMigracion.MigrarTodasLasResidencias()` desde el editor VBA (F5)
- [ ] Ejecutar `modMigracion.VerificarMigracion()` para comprobar integridad
- [ ] Los registros se migran de las hojas RESIDENCIA GIJÓN/SOTO/OVIEDO + LISTA NEGRA + LOG

### Fase 4: Adaptar los 15 módulos VBA existentes (según Guia-Modificacion-Modulos-VBA.md)
Los módulos a modificar con el patrón de doble escritura son:

| # | Módulo | Tipo de cambio | Complejidad |
|---|--------|---------------|-------------|
| 1 | `FechasPeticion.bas` | Sustituir bloque de grabación | ⭐⭐ Media |
| 2 | `EvitarDuplicidadSolicitudesGyS.bas` | Sustituir bucle For por consulta SQL | ⭐ Baja |
| 3 | `AsignarNumFactura.bas` | Sustituir función de bloques por contador atómico | ⭐⭐ Media |
| 4 | `FacturacionMesGIJON.bas` | Añadir línea de UPDATE tras escritura en celda | ⭐ Baja |
| 5 | `FacturacionMesOVIEDO.bas` | Ídem | ⭐ Baja |
| 6 | `FacturacionMesSOTO.bas` | Ídem | ⭐ Baja |
| 7 | `MarcarSiPagadosEnResidencia.bas` | Añadir línea de UPDATE | ⭐ Baja |
| 8 | `ModuloCalendarioGijon.bas` | Añadir 1 línea de sincronización al inicio | ⭐ Baja |
| 9 | `ModuloCalendarioOviedo.bas` | Ídem | ⭐ Baja |
| 10 | `ModuloCalendarioSoto.bas` | Ídem | ⭐ Baja |
| 11 | `BusquedaDNIResidencias.bas` | Sustituir Find por consulta SQL | ⭐ Baja |
| 12 | `BusquedaOrdenNombreFactura.bas` | Ídem | ⭐ Baja |
| 13 | `ModuloLOG.bas` | Sustituir 4 líneas por 1 | ⭐ Baja |
| 14 | `ModListaNegra.bas` | Sustituir Find por consulta SQL | ⭐ Baja |
| 15 | `ReevaluacionSolicitudes.bas` | Añadir sincronización + UPDATE en bucle | ⭐⭐ Media |

Además:
- [ ] Añadir eventos `Worksheet_Change` con propagación a Access (módulo 11 de la guía)
- [ ] Modificar `Workbook_Open` para sincronización al arranque (módulo 12 de la guía)

### Fase 5: Pruebas y despliegue
- [ ] Probar con 2 copias de Excel abiertas simultáneamente contra la misma BD
- [ ] Verificar que las inserciones no generan duplicados de Nº Orden
- [ ] Verificar sincronización de calendario entre operadores
- [ ] Desplegar a los 6 PCs de producción

---

## 🏗️ Arquitectura de la solución

```
┌─────────────────────────────────────────────────────────┐
│                    UNIDAD H:\ (RED)                     │
│                                                         │
│  H:\ResidenciaBD\Residencia_BE.accdb                    │
│  ┌─────────────────────────────────────────────────┐    │
│  │  Ordenes (NumOrden AUTOINCREMENT)               │    │
│  │  ContadorFacturas (por residencia + ejercicio)  │    │
│  │  LogActividad                                    │    │
│  │  ListaNegra                                      │    │
│  └─────────────────────────────────────────────────┘    │
│                         ▲                               │
│                    ADO / OLEDB                           │
│                    (vida corta)                          │
└─────────────────────────────────────────────────────────┘
          ▲         ▲         ▲         ▲
          │         │         │         │
     ┌────┴──┐ ┌───┴───┐ ┌───┴───┐ ┌───┴───┐
     │ PC #1 │ │ PC #2 │ │ PC #3 │ │ PC #4 │  ... (hasta 6)
     │ Excel │ │ Excel │ │ Excel │ │ Excel │
     │ local │ │ local │ │ local │ │ local │
     └───────┘ └───────┘ └───────┘ └───────┘
```

---

## 💬 Frase para retomar

Cuando vuelvas, dile al asistente:

> **"Quiero continuar con la implementación BDAS Multiusuario (Excel + Access en H:\). Lee el archivo `H:\ResidenciaApp\bdas-multiusuario\scripts\ESTADO-IMPLEMENTACION-BDAS-MULTIUSUARIO.md` para ver dónde quedamos."**

---
