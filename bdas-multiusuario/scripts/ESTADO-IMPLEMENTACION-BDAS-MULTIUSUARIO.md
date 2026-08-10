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

### Fase 2: Importar módulos VBA en el Excel BDAS
- [ ] Abrir el Excel maestro → ALT+F11
- [ ] Importar `modDatabase.bas`
- [ ] Importar `modMigracion.bas`
- [ ] Crear UserForm `frmLogin` (siguiendo instrucciones en `frmLogin.frm`)
- [ ] Crear UserForm `frmSelectorResidencia` (siguiendo instrucciones en `frmSelectorResidencia.frm`)
- [ ] Pegar código de `ThisWorkbook_Events.bas` en ThisWorkbook

### Fase 3: Migración de datos históricos
- [ ] Ejecutar `modMigracion.MigrarTodasLasResidencias()` desde el editor VBA (F5)
- [ ] Ejecutar `modMigracion.VerificarMigracion()` para comprobar integridad
- [ ] Los registros se migran de las hojas RESIDENCIA GIJÓN/SOTO/OVIEDO + LISTA NEGRA + LOG

### Fase 4: Adaptar los 15 módulos VBA existentes (según Guia-Modificacion-Modulos-VBA.md)
Los módulos a modificar (ya NO doble escritura, solo Access + refrescar caché):

| # | Módulo | Tipo de cambio | Complejidad |
|---|--------|---------------|-------------|
| 1 | `FechasPeticion.bas` | Solo `InsertarOrdenAtomica()` + `RefrescarCacheVisual()` | ⭐⭐ Media |
| 2 | `EvitarDuplicidadSolicitudesGyS.bas` | Consulta SQL (sin cambios del plan original) | ⭐ Baja |
| 3 | `AsignarNumFactura.bas` | Contador atómico (sin cambios) | ⭐⭐ Media |
| 4 | `FacturacionMesGIJON.bas` | Solo `ActualizarOrden()` + `RefrescarCacheVisual()` | ⭐ Baja |
| 5 | `FacturacionMesOVIEDO.bas` | Ídem | ⭐ Baja |
| 6 | `FacturacionMesSOTO.bas` | Ídem | ⭐ Baja |
| 7 | `MarcarSiPagadosEnResidencia.bas` | Solo `ActualizarOrden()` + `RefrescarCacheVisual()` | ⭐ Baja |
| 8 | `ModuloCalendarioGijon.bas` | Sin cambios (ya sincroniza desde Access) | ⭐ Baja |
| 9 | `ModuloCalendarioOviedo.bas` | Ídem | ⭐ Baja |
| 10 | `ModuloCalendarioSoto.bas` | Ídem | ⭐ Baja |
| 11 | `BusquedaDNIResidencias.bas` | Consulta SQL (sin cambios) | ⭐ Baja |
| 12 | `BusquedaOrdenNombreFactura.bas` | Ídem | ⭐ Baja |
| 13 | `ModuloLOG.bas` | `InsertarLog()` (sin cambios) | ⭐ Baja |
| 14 | `ModListaNegra.bas` | Consulta SQL (sin cambios) | ⭐ Baja |
| 15 | `ReevaluacionSolicitudes.bas` | Solo `ActualizarOrden()` en bucle + `RefrescarCacheVisual()` | ⭐⭐ Media |

### Fase 5: Pruebas y despliegue
- [ ] Probar con 2 PCs abriendo el .xlsm simultáneamente
- [ ] Verificar login manual + selector de residencia
- [ ] Verificar que NO pide guardar al cerrar
- [ ] Verificar que inserciones de un usuario son visibles tras refrescar en otro
- [ ] Verificar que no hay duplicados de Nº Orden
- [ ] Copiar `BDAS_v16.5.5.xlsm` a `H:\ResidenciaBD\`
- [ ] Crear accesos directos en los 15 PCs

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
