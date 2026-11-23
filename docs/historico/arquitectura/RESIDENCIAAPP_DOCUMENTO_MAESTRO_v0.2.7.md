# RESIDENCIAAPP — DOCUMENTO MAESTRO DE ARQUITECTURA Y REQUISITOS
## VERSIÓN 0.2.7 — CANDIDATA A REVISIÓN ARQUITECTÓNICA INDEPENDIENTE

**Fecha de emisión:** 16 de noviembre de 2026  
**Estado:** CANDIDATA A REVISIÓN ARQUITECTÓNICA INDEPENDIENTE  
**Ámbito:** ResidenciaApp Multiusuario (Excel VBA + Access OLEDB)  
**SHA-256 Front-End Producción (`BDAS_Multiusuario.xlsm`):** `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`  
**SHA-256 Back-End Producción (`Residencia_BE.accdb`):** `ac2fb7240360b3771eca60f61ca8cf32fffeaf78e88b09d783c169a228bfe5cf`  
**Base Documental:** Inspección As-Built Órdenes 0, 0A, 0B, 0C, Informes 1A/1B, Cierre Técnico Dirigido Orden 1C, Rectificación Final Orden 1D, Rectificación Orden 1E, Rectificación Documental Final Orden 1F, Rectificación Documental Orden 1G y verificaciones de identidad/extracción del módulo de persistencia y precisión de CodeName/EstadoPago de 16/11/2026 (Órdenes 1H y 1J).

---

### ÍNDICE GENERAL
1. [Principio Fundamental: Las Tres Realidades Separadas](#1-principio-fundamental-las-tres-realidades-separadas)
2. [Bloque I — Arquitectura As-Built de Producción](#2-bloque-i--arquitectura-as-built-de-producción)
   - 2.1. Topología Física y Despliegue de Red
   - 2.2. Capas del Sistema, Desacoplamiento y Orquestación Procedimental
   - 2.3. Interfaz de Usuario y Componentes VBA (173 en Binario vs 198 en Repositorio)
   - 2.4. Estructura de Cuadrícula y Columnas por Residencia (GIJÓN, SOTO, OVIEDO)
   - 2.5. Autenticación, Permisos, Token de Instancia (CoCreateGuid) y Roles (ADMIN y OPERADOR)
   - 2.6. Persistencia y Operaciones ADODB Directas (32 Procedimientos en modDatabase)
   - 2.7. Gestión Operativa del Nº ORDEN en Producción
   - 2.8. Gestión de Reservas, Huéspedes y Facturación
   - 2.9. Gestión de Tarifas, Precios y Turnos (Hoja28 / Turnos-Precios y TARIFAS UNIDAD)
   - 2.10. Calendarios, Habitaciones y Bloqueos
   - 2.11. Subsistema de Contingencia Desconectada (SICO-N2) y Registro Local WAL
   - 2.12. Ciclo de Vida de Datos, Archivado y Purgas (5 Procedimientos Reales)
   - 2.13. Limitaciones Funcionales y Defectos Estructurales de Producción
3. [Bloque II — Arquitectura Objetivo Aprobada](#3-bloque-ii--arquitectura-objetivo-aprobada)
   - 3.1. Marco Arquitectónico V5.2 y Principios Rectores
   - 3.2. Motor Transaccional Especializado: `clsDatabaseEngine` (Estado Real de Laboratorio)
   - 3.3. Clave de Capacidad (CapKey de 256 bits), Validación y Escritura Protegida
   - 3.4. Control de Concurrencia Optimista y Lease Lock (120 Segundos)
   - 3.5. Asignación Atómica de Nº ORDEN y Algoritmo de Búsqueda de Huecos
   - 3.6. Segregación Canónica de Rangos Numéricos (1-899999 Oficial, 900001-999999 Provisional)
   - 3.7. Reconciliación FIFO Segura
   - 3.8. Persistencia Granular de Tarifas (TAR-10) vs Persistencia Base de Tarifas
   - 3.9. Estado de Aprobación vs. Propuestas Pendientes
4. [Bloque III — Catálogo de Defectos, Riesgos y Evolución Pendiente](#4-bloque-iii--catálogo-de-defectos-riesgos-y-evolución-pendiente)
   - 4.1. Fichas Técnicas de los 18 Defectos Estables
5. [Matriz Maestra de Requisitos del Sistema](#5-matriz-maestra-de-requisitos-del-sistema)
   - 5.1. Categoría A: Requisitos del Negocio (Trazabilidad y Certeza Documental)
   - 5.2. Categoría B: Requisitos Incorporados en Fases de Desarrollo
   - 5.3. Categoría C: Requisitos Aprobados para Arquitectura Objetivo
   - 5.4. Categoría D: Mejoras Propuestas o Aplazadas (Depuradas de Propuestas Externas)
6. [Modelo Físico de Datos Detallado (11 Tablas de Usuario)](#6-modelo-físico-de-datos-detallado-11-tablas-de-usuario)
   - Especificación Completa de Campos, Claves, Índices Reales y Ausencia de Claves Foráneas
7. [Inventario de Componentes y Trazabilidad de Fuentes](#7-inventario-de-componentes-y-trazabilidad-de-fuentes)
   - 7.1. Cifras Verificadas y Desglose de los 173 Componentes del Binario
   - 7.2. Catálogo Acreditado de los 8 Archivos Huérfanos / Redundantes en Repositorio
   - 7.3. Recurso RibbonX (`customUI/customUI.xml` en Paquete XLSM)
   - 7.4. Trazabilidad Línea a Línea y Hashes Canónicos
8. [Reglas de Certeza, Validación y Cuestiones Jurídicas](#8-reglas-de-certeza-validación-y-cuestiones-jurídicas)
   - 8.1. Taxonomía de Niveles de Certeza Técnica
   - 8.2. Cuestiones Jurídicas y Periodos de Retención (RD 933/2021)
9. [Incertidumbres Técnicas que Permanecen Abiertas](#9-incertidumbres-técnicas-que-permanecen-abiertas)
10. [Reglas de Gobernanza del Documento Maestro y Registro de Versiones](#10-reglas-de-gobernanza-del-documento-maestro-y-registro-de-versiones)
11. [Control de Cambios Detallado (Versión 0.2 a Versión 0.2.7)](#11-control-de-cambios-detallado-versión-02-a-versión-027)
    - 11.1. Rectificaciones Incorporadas de Versión 0.2 a Versión 0.2.1
    - 11.2. Rectificaciones Incorporadas de Versión 0.2.1 a Versión 0.2.2 (Orden 1C)
    - 11.3. Rectificaciones Incorporadas de Versión 0.2.2 a Versión 0.2.3 (Orden 1D)
    - 11.4. Rectificaciones Incorporadas de Versión 0.2.3 a Versión 0.2.4 (Orden 1E)
    - 11.5. Rectificaciones Incorporadas de Versión 0.2.4 a Versión 0.2.5 (Orden 1F)
    - 11.6. Rectificaciones Incorporadas de Versión 0.2.5 a Versión 0.2.6 (Orden 1G)
    - 11.7. Rectificaciones Incorporadas de Versión 0.2.6 a Versión 0.2.7 (Orden 1H)
    - 11.8. Precisiones Complementarias a la Versión 0.2.7 (Orden 1J)
- [ANEXO A — Inventario Exhaustivo de Operaciones ADODB de Producción (32 Procedimientos Directos y 2 Orquestadores Compuestos)](#anexo-a--inventario-exhaustivo-de-operaciones-adodb-de-producción-32-procedimientos-directos-y-2-orquestadores-compuestos)

---

### 1. PRINCIPIO FUNDAMENTAL: LAS TRES REALIDADES SEPARADAS

Este Documento Maestro estructura el conocimiento arquitectónico de ResidenciaApp en tres realidades estrictamente delimitadas para evitar confusiones entre el código activo en producción, las soluciones arquitectónicas formalmente aprobadas en laboratorio y los problemas pendientes de resolución:

```
┌────────────────────────────────────────────────────────────────────────┐
│ REALIDAD 1: BLOQUE I — ARQUITECTURA AS-BUILT DE PRODUCCIÓN             │
│ Describe exclusivamente lo que existe y se ejecuta en el libro activo  │
│ BDAS_Multiusuario.xlsm (SHA-256: b9117659...) y su backend Access      │
│ Residencia_BE.accdb (SHA-256: ac2fb724...). Es la verdad fáctica       │
│ operativa con sus virtudes y sus defectos reales.                      │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│ REALIDAD 2: BLOQUE II — ARQUITECTURA OBJETIVO APROBADA                 │
│ Soluciones técnicas estructuradas en dos niveles de madurez formal:    │
│ 1) Código integrado en laboratorio (Fase 4): clsDatabaseEngine V5.2   │
│    y CapKey V5.2 (en binario BDAS_Multiusuario_Fase4.xlsm).            │
│ 2) Arquitectura objetivo aprobada pendiente de implementación:         │
│    Búsqueda de huecos ORD-10 Rev. 5, Reconciliación FIFO Rev. 5,       │
│    Segregación canónica de rangos y Persistencia TAR-10 (diseño).      │
│ (Nota: El Lease Lock base ya es operativo en producción; en laboratorio│
│ se implementó su encapsulación en clsDatabaseEngine, clasificada en    │
│ Certeza 4 por falta de evidencia de pruebas unitarias en repositorio). │
└───────────────────────────────────┬────────────────────────────────────┘
                                    │
                                    ▼
┌────────────────────────────────────────────────────────────────────────┐
│ REALIDAD 3: BLOQUE III — DEFECTOS, RIESGOS Y PROPUESTAS PENDIENTES     │
│ Inventario versionado de 18 defectos estables detectados, riesgos de │
│ concurrencia, colisiones y propuestas en fase de diseño o corrección. │
└────────────────────────────────────────────────────────────────────────┘
```

Ninguna afirmación de diseño o deseo de evolución debe atribuirse al Bloque I si no se encuentra físicamente inyectada y operativa en el binario activo. Asimismo, dentro del Bloque II se distinguen con rigor los componentes con código presente e inspeccionado en el binario de laboratorio Fase 4 (como la clase `clsDatabaseEngine` y la `CapKey` V5.2) de aquellas soluciones que cuentan con diseño formal aprobado pero cuya implementación o inyección en código ejecutable aún está pendiente (como la búsqueda de huecos ORD-10 Rev. 5, la reconciliación FIFO Rev. 5, la segregación canónica de rangos numéricos y la persistencia granular de tarifas TAR-10). En relación con `clsDatabaseEngine` V5.2 y `CapKey` V5.2, su código ha sido extraído e inspeccionado documentalmente en el binario `BDAS_Multiusuario_Fase4.xlsm`, pero su homologación funcional completa no está acreditada con registros reproducibles suficientes en la documentación permanente examinada, permaneciendo ambos componentes clasificados en Certeza 4 (Incertidumbre / Pendiente de Evidencia de Pruebas). Esta incertidumbre afecta estrictamente a la acreditación formal de su validación funcional mediante registros de prueba en repositorio, sin cuestionar la existencia física del código extraído ni declarar que el componente sea erróneo o inoperativo.

---

### 2. BLOQUE I — ARQUITECTURA AS-BUILT DE PRODUCCIÓN

#### 2.1. Topología Física y Despliegue de Red
ResidenciaApp es una aplicación multiusuario cliente-servidor basada en archivos sobre red de área local (protocolo SMB/CIFS):
- **Front-end:** Libro habilitado para macros `BDAS_Multiusuario.xlsm` (SHA-256: `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`), ejecutado localmente en instancias independientes de Microsoft Excel por los operadores en puestos clientes de la red.
- **Back-end de Producción Activa:** Base de datos relacional de Microsoft Access `H:\ResidenciaBD\Residencia_BE.accdb` (SHA-256: `ac2fb7240360b3771eca60f61ca8cf32fffeaf78e88b09d783c169a228bfe5cf`), ubicada en un recurso compartido de red accesible mediante el proveedor OLEDB `Microsoft.ACE.OLEDB.12.0`.
- **Mecanismo Exacto de Resolución de Ruta:**
  La propiedad `Property Get DB_PATH()` y la función `ObtenerRutaBaseDatos()` en `modDatabase.bas` (líneas 28 a 111) resuelven la ubicación del archivo backend ejecutando la siguiente cascada determinista:
  1. **Comprobación 1 (Directorio Local del Libro):** Si `ThisWorkbook.Path` no contiene la subcadena `"C:\RESIDENCIAAPP"`, comprueba si existe el archivo `ThisWorkbook.Path & "\Residencia_BE.accdb"`. Si existe, retorna dicha ruta local.
  2. **Comprobación 2 (Ruta de Preproducción en Red):** Si la comprobación 1 no se cumple, evalúa la existencia de la constante `DB_PATH_PREPROD`:
     ```vb
     Public Const DB_PATH_PREPROD As String = "H:\GIJ1\AYUDANTIA\01 RESIDENCIAS\BDAS Asturias\Residencia_BE.accdb"
     ```
     Si dicho archivo existe en la red, retorna `DB_PATH_PREPROD`.
  3. **Comprobación 3 (Ruta de Desarrollo / Red Local):** Si la comprobación 2 no se cumple, evalúa la constante `DB_PATH_DEV`:
     ```vb
     Public Const DB_PATH_DEV As String = "H:\ResidenciaBD\Residencia_BE.accdb"
     ```
     Si dicho archivo existe, retorna `DB_PATH_DEV`.
  4. **Comprobación 4 (Fallback Estructural):** Si ninguna de las rutas anteriores existe, evalúa si existe el directorio `H:\ResidenciaApp`. Si existe, devuelve `"H:\ResidenciaBD\Residencia_BE.accdb"`, y en caso contrario devuelve `"C:\ResidenciaBD\Residencia_BE.accdb"`.
- **Comportamiento Verificado en el Entorno Activo:**  
  Dado que el libro se ejecuta desde `H:\ResidenciaApp\bdas-multiusuario`:
  - La comprobación 1 busca `H:\ResidenciaApp\bdas-multiusuario\Residencia_BE.accdb`, la cual no existe.
  - La comprobación 2 busca `H:\GIJ1\AYUDANTIA\01 RESIDENCIAS\BDAS Asturias\Residencia_BE.accdb`, la cual no existe en la máquina operativa actual.
  - La comprobación 3 evalúa `H:\ResidenciaBD\Residencia_BE.accdb`, cuya existencia física y validez OLEDB han sido confirmadas plenamente.  
  Por tanto, **la ruta activa resuelta en producción es exclusivamente `H:\ResidenciaBD\Residencia_BE.accdb`**.
- **Delimitación de Rutas en Otros Entornos del Proyecto:**
  - **Entorno de Laboratorio / Pruebas Automatizadas:** `H:\ResidenciaApp\bdas-multiusuario\tests\harness\Residencia_BE.accdb` (backend aislado utilizado exclusivamente por los arneses de test).
  - **Réplica Golden Master (Inmutable):** `H:\ResidenciaBD\Plantilla\GoldenMaster\Residencia_BE_GOLDEN_MASTER.accdb` (hito 2026-10-24).
  - **Réplica Versión Referencia (Inmutable):** `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` (hito 2026-11-10).

#### 2.2. Capas del Sistema, Desacoplamiento y Orquestación Procedimental
El sistema presenta una arquitectura de dos niveles desacoplados físicamente:
1. **Capa de Presentación y Lógica de Negocio (Microsoft Excel / VBA):**  
   Gestiona la interfaz gráfica mediante hojas de cálculo estructuradas y formularios UserForm. Alberga la lógica de validación de entradas, cálculo de tarifas, gestión de calendarios, renderizado visual y control de interfaz.
2. **Capa de Almacenamiento Relacional (Microsoft Access / ACE OLEDB):**  
   Almacena el estado persistente centralizado en 11 tablas de usuario relacionales.
- **Límites de Transaccionalidad y Ausencia de Transacciones Distribuidas:**  
  Debe destacarse con total rigor técnico que **no existe una transacción atómica distribuida global que abarque simultáneamente las celdas de Excel y las tablas de Access**. Microsoft Excel (un modelo de objetos COM de hoja de cálculo) y Microsoft Access (un motor de base de datos relacional accedido mediante OLEDB) son subsistemas heterogéneos y físicamente independientes. En el entorno VBA de Excel no participa ningún Coordinador de Transacciones Distribuidas (MSDTC o protocolo Two-Phase Commit).  
  Por tanto:
  - La atomicidad formal en el sistema queda **estrictamente circunscrita a las transacciones internas de Access** ejecutadas bajo `cn.BeginTrans`, `cn.CommitTrans` y `cn.RollbackTrans`.
  - La coherencia entre el contenido visible en las hojas de Excel y los registros de Access se mantiene exclusivamente mediante **orquestación procedimental secuencial en VBA** (verificación previa de precondiciones, ejecución del comando SQL y posterior actualización o compensación de las celdas de la cuadrícula).

#### 2.3. Interfaz de Usuario y Componentes VBA (173 en Binario vs 198 en Repositorio)
La auditoría independiente y la extracción física de componentes macro del libro de producción `BDAS_Multiusuario.xlsm` (SHA-256: `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`) confirma de forma concluyente la existencia de **exactamente 173 componentes VBA en el binario activo**:
- **106 Módulos Estándar (`.bas`):** Procedimientos de base de datos (`modDatabase`), seguridad local (`ModuloSeguridadLocal`), facturación, cálculo y calendarios.
- **49 Módulos de Hoja (`.cls`):** Módulos asociados a las hojas de cálculo operativas, maestras de residencia (`Hoja2`, `Hoja20`, `Hoja24`), tarifas (`Hoja28`), calendarios y hojas ocultas del sistema (`SYS_COLA_SYNC`, `SYS_CONFIG`).
- **1 Módulo `ThisWorkbook` (`.cls`):** Controlador central de eventos de apertura, cierre, cambio de hoja y menús RibbonX.
- **17 Formularios de Usuario (`.frm`):** Interfaz gráfica modal y no modal (`frmLogin`, `frmGestionListaNegra`, `frmBloqueosHabitaciones`, `frmPanelRGPD`, etc.).
- **Total en Binario Activo:** **173 componentes VBA** (106 + 49 + 1 + 17).

> [!IMPORTANT] Diferenciación entre Binario Activo y Archivos Exportados en Disco
> En la carpeta del repositorio `vba-modules/` existen físicamente **198 archivos exportados** (113 `.bas`, 51 `.cls`, 17 `.frm`, 17 `.frx`). La diferencia de 25 archivos respecto a los 173 componentes activos del binario se compone de:
> - **17 archivos de recursos de formularios (`.frx`)**: Recursos binarios de controles de los 17 UserForms.
> - **8 archivos huérfanos / redundantes**: Archivos `.bas` y `.cls` que no forman parte del proyecto binario compilado (ver catálogo detallado en Sección 7.2).
> Asimismo, se deja constancia de que la afirmación de ausencia de errores de compilación debe sustentarse en pruebas dinámicas independientes en tiempo de ejecución de Excel.

#### 2.4. Estructura de Cuadrícula y Columnas por Residencia (GIJÓN, SOTO, OVIEDO)
Las hojas maestras de órdenes no poseen una estructura de columnas idéntica. Cada residencia presenta particularidades funcionales en su cuadrícula de datos, reflejadas en sus módulos de hoja (`Hoja2.cls`, `Hoja20.cls` y `Hoja24.cls`), en la función de mapeo de eventos `MapearColumnaACampoBD` (`ModuloSeguridadLocal.bas:683-741`) y en el procedimiento de volcado `SincronizarFilaHaciaAccess` (`modDatabase.bas:1737-1993`). La correspondencia definitiva entre nombre visible, CodeName VBA y parte XML del libro se recoge a continuación.

##### Matriz Canónica de Correspondencia de Columnas y Persistencia (Cinco Dimensiones)
En cumplimiento de la Orden 1F, la siguiente tabla desglosa de forma unificada las cinco dimensiones requeridas, contrastadas directamente con la estructura XML de las hojas (`sheet2.xml`, `sheet23.xml`, `sheet34.xml`), los encabezados físicos de la Fila 1 y el código VBA de persistencia. La verificación de 16/11/2026 acredita además que el `modDatabase.bas` del repositorio y el módulo `VBA/modDatabase` extraído del binario activo son idénticos byte a byte (298.794 bytes; SHA-256 `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47`):

| Col. | GIJÓN: Encabezado Real | SOTO: Encabezado Real | OVIEDO: Encabezado Real | ¿Leída en Sincronizar? | Asignación en Access (SincronizarFilaHaciaAccess) | Persistencia en Access mediante otras rutinas independientes (si está acreditada) |
| :---: | :--- | :--- | :--- | :---: | :--- | :--- |
| **1 (A)** | Nº ORDEN | Nº ORDEN | Nº ORDEN | **Sí** | NumOrden (identificador en WHERE; asignado en AddNew) | `SincronizarLoteImportacionBD`: identifica registros existentes y asigna `NumOrden` en `AddNew` cuando el lote crea una fila; además, `IntentarReservarOrdenBD`, `ReservarSiguienteOrdenLibreBD`, `RenumerarOrdenExistenteBD` y `AsegurarOrdenExiste` operan también sobre `Ordenes.NumOrden` (según el inventario previo; no revalidados por la Orden 1H) |
| **2 (B)** | FECHA PETICION | FECHA PETICION | FECHA PETICION | **Sí** | FechaPeticion (DateTime) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `FechaPeticion` en la rutina por lotes |
| **3 (C)** | NÚM FACT. | NÚM FACT. | NÚM FACT | **Sí** | NumFactura (Long) en las 3 sedes (preservada si vacío) | `SincronizarLoteImportacionBD`: asigna `NumFactura`; preserva el valor previo si la celda viene vacía en una fila existente |
| **4 (D)** | FINALIDAD | FINALIDAD | FINALIDAD | **Sí** | Finalidad (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Finalidad` en la rutina por lotes |
| **5 (E)** | EMPLEO | EMPLEO | EMPLEO | **Sí** | Empleo (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Empleo` en la rutina por lotes |
| **6 (F)** | SITUACION | SITUACION | SITUACION | **Sí** | Situacion (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Situacion` en la rutina por lotes |
| **7 (G)** | EVALUACIÓN | EVALUACIÓN | EVALUACIÓN | **Sí** | Evaluacion (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Evaluacion` en la rutina por lotes |
| **8 (H)** | COMISIÓN | **TURNO** | COMISIÓN | **Sí** | SOTO: Turno / GIJÓN y OVIEDO: Comision | `SincronizarLoteImportacionBD`: SOTO → `Turno`; GIJÓN y OVIEDO → `Comision` |
| **9 (I)** | DNI | DNI | DNI | **Sí** | DNI (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `DNI` en la rutina por lotes |
| **10 (J)**| RANGO | RANGO | RANGO | **Sí** | Rango (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Rango` en la rutina por lotes |
| **11 (K)**| Nombre | Nombre | Nombre | **Sí** | Nombre (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Nombre` en la rutina por lotes |
| **12 (L)**| ENTRADA | ENTRADA | ENTRADA | **Sí** | FechaEntrada (DateTime) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `FechaEntrada` en la rutina por lotes |
| **13 (M)**| SALIDA | SALIDA | SALIDA | **Sí** | FechaSalida (DateTime) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `FechaSalida` en la rutina por lotes |
| **14 (N)**| DIAS USO | DIAS USO | DIAS USO | **Sí** | DiasUso (Long) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `DiasUso` en la rutina por lotes |
| **15 (O)**| PAX | PAX | PAX | **Sí** | PAX (Long) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `PAX` en la rutina por lotes |
| **16 (P)**| RESOLUCION | RESOLUCION | RESOLUCION | **Sí** | Resolucion (Text) en las 3 sedes | `SincronizarLoteImportacionBD`: asigna `Resolucion` en la rutina por lotes |
| **17 (Q)**| HAB. IND. | **PRECIO APTO.** | HAB. IND. | **Sí** | GIJÓN y OVIEDO: NumHabIndividuales / SOTO: Ninguno (no asignado) | `SincronizarLoteImportacionBD`: GIJÓN y OVIEDO: sí, `NumHabIndividuales`; SOTO: no, se omite expresamente |
| **18 (R)**| HAB. DOBLE| **SUPLE. OCUPAN.**| HAB. DOBLE | **Sí** | GIJÓN y OVIEDO: NumHabDobles / SOTO: Ninguno (no asignado) | `SincronizarLoteImportacionBD`: GIJÓN y OVIEDO: sí, `NumHabDobles`; SOTO: no, se omite expresamente |
| **19 (S)**| NÚM HAB. | APARTAMENTO | **CAMA SUPLE.** | **Sí** | GIJÓN y SOTO: HabitacionesAsignadas / OVIEDO: CamaSuple | `SincronizarLoteImportacionBD`: asigna el campo específico de cada sede (`HabitacionesAsignadas` o `CamaSuple`) |
| **20 (T)**| DTO. FAM. NUM.| DTO. FAM. NUM. | **NÚM HAB.** | **Sí** | GIJÓN y SOTO: DtoFamNum / OVIEDO: HabitacionesAsignadas | `SincronizarLoteImportacionBD`: asigna `DtoFamNum` en GIJÓN/SOTO y `HabitacionesAsignadas` en OVIEDO |
| **21 (U)**| IMPORTE | IMPORTE | **DTO. FAM. NUM.**| **Sí** | GIJÓN y SOTO: Importe (Double) / OVIEDO: DtoFamNum | `SincronizarLoteImportacionBD`: asigna `Importe` en GIJÓN/SOTO y `DtoFamNum` en OVIEDO |
| **22 (V)**| TELEFONO | TELEFONO | **IMPORTE** | **Sí** | GIJÓN y SOTO: Telefono / OVIEDO: Importe (Double) | `SincronizarLoteImportacionBD`: asigna `Telefono` en GIJÓN/SOTO y `Importe` en OVIEDO |
| **23 (W)**| DIRECCIÓN | DIRECCIÓN | **TELEFONO** | **Sí** | GIJÓN y SOTO: Direccion / OVIEDO: Telefono | `SincronizarLoteImportacionBD`: asigna `Direccion` en GIJÓN/SOTO y `Telefono` en OVIEDO |
| **24 (X)**| CP | CP | **DIRECCIÓN** | **Sí** | GIJÓN y SOTO: CodigoPostal / OVIEDO: Direccion | `SincronizarLoteImportacionBD`: asigna `CodigoPostal` en GIJÓN/SOTO y `Direccion` en OVIEDO |
| **25 (Y)**| POBLACIÓN | POBLACIÓN | **CP** | **Sí** | GIJÓN y SOTO: Poblacion / OVIEDO: CodigoPostal | `SincronizarLoteImportacionBD`: asigna `Poblacion` en GIJÓN/SOTO y `CodigoPostal` en OVIEDO |
| **26 (Z)**| PROVINCIA | PROVINCIA | **POBLACIÓN** | **Sí** | GIJÓN y SOTO: Provincia / OVIEDO: Poblacion | `SincronizarLoteImportacionBD`: asigna `Provincia` en GIJÓN/SOTO y `Poblacion` en OVIEDO |
| **27 (AA)**| PAGADO | PAGADO | **PROVINCIA** | **Sí** | GIJÓN y SOTO: EstadoPago / OVIEDO: Provincia | `SincronizarLoteImportacionBD`: GIJÓN/SOTO: sí, `EstadoPago`; OVIEDO: sí, `Provincia` |
| **28 (AB)**| GRABACIÓN SOLICITUD | GRABACIÓN SOLICITUD | **PAGADO** | **Sí** | GIJÓN y SOTO: Ninguno (no asignado) / OVIEDO: EstadoPago | `SincronizarLoteImportacionBD`: GIJÓN/SOTO: no, no asigna AB; OVIEDO: sí, asigna `EstadoPago` |
| **29 (AC)**| COPIAR TEXTO EMAIL | COPIAR TEXTO EMAIL | REPARTIR COSTE  HAB.| **No** | Ninguno (columna técnica fuera de la matriz leída) | No: AC queda fuera del rango A:AB leído por la rutina |
| **30 (AD)**| REPARTIR COSTE  HAB.| REPARTIR COSTE  APTO.| TAG | **No** | Ninguno (columna técnica fuera de la matriz leída) | No: AD queda fuera del rango A:AB leído por la rutina |
| **31 (AE)**| *(No existe)* | *(No existe)* | R | **No** | Ninguno (columna técnica fuera de la matriz leída) | No: AE queda fuera del rango A:AB leído por la rutina |
| **32 (AF)**| *(No existe)* | *(No existe)* | GRABACIÓN SOLICITUD | **No** | Ninguno (columna técnica fuera de la matriz leída) | No: AF queda fuera del rango A:AB leído por la rutina |
| **33 (AG)**| *(No existe)* | *(No existe)* | COPIAR TEXTO EMAIL | **No** | Ninguno (columna técnica fuera de la matriz leída) | No: AG queda fuera del rango A:AB leído por la rutina |

**Alcance de las dos rutinas de escritura:** `SincronizarFilaHaciaAccess` realiza la sincronización individual de una fila y adquiere previamente el Lease Lock mediante un `UPDATE` condicional. `SincronizarLoteImportacionBD` es una segunda ruta de persistencia, destinada a importar un conjunto de filas: recorre el Recordset de la residencia, crea o actualiza cada registro y aplica el mapeo de la matriz. Por tanto, no es correcto afirmar que la factura o el estado de pago carezcan de persistencia distinta de la sincronización individual. No existen, en cambio, procedimientos acreditados con los nombres `ActualizarFacturaBD` o `ActualizarEstadoPagoBD`; ambos valores los persiste la rutina por lotes indicada.

La rutina por lotes conserva `NumFactura` y `EstadoPago` ya existentes cuando la celda correspondiente llega vacía en una fila preexistente. Para el resto de campos no consta esa misma protección general frente a celdas vacías. La rutina por lotes no establece por sí misma `EstadoPago = "PENDIENTE"` como valor predeterminado: para registros nuevos llama a `AsignarCampoTexto` con el contenido de la celda.

**Precisión sobre la ruta de reserva:** `IntentarReservarOrdenBD`, cuando el número solicitado está libre y ejecuta la rama `INSERT`, incluye expresamente `EstadoPago` en la lista de campos y escribe el literal `'PENDIENTE'` (líneas 1358–1369 del módulo verificado). Esta constatación acredita el valor inicial en esa ruta concreta de reserva; no demuestra que el campo Access tenga configurado un `DefaultValue` general. La rama `forzarReserva=True` actualiza un registro existente y no equivale a esa inserción inicial.

**Límite de concurrencia de la importación masiva:** aunque `SincronizarLoteImportacionBD` ejecuta `cn.BeginTrans` y confirma al final del lote, no contiene una adquisición condicional por expediente ni consulta si el Lease Lock pertenece a otra instancia. En registros existentes, actualiza los datos y sustituye `UsuarioModificacion` por el token actual. La transacción de Access no equivale a la protección lógica de concurrencia del Lease Lock; esta ruta puede eludir el bloqueo que aplica `SincronizarFilaHaciaAccess`. Véase `DEF-FIFO-07`.

**Identificación definitiva de las hojas maestras (CodeName VBA):** el volcado estructural del libro activo, cruzando la relación de hojas del paquete con los CodeName del proyecto VBA, confirma la siguiente correspondencia:

| Nombre visible de la hoja | CodeName VBA | Parte de hoja en el paquete XLSX/XLSM |
|---|---|---|
| RESIDENCIA GIJÓN | `Hoja2` | `xl/worksheets/sheet2.xml` |
| RESIDENCIA SOTO | `Hoja20` | `xl/worksheets/sheet23.xml` |
| RESIDENCIA OVIEDO | `Hoja24` | `xl/worksheets/sheet34.xml` |
| Calendario GIJÓN | `Hoja1` | No necesario para la matriz de persistencia |
| Calendario OVIEDO | `Hoja23` | No necesario para la matriz de persistencia |
| TARIFAS UNIDAD | `Hoja34` | No necesario para la matriz de persistencia |

La extracción utilizada en la Orden 1H confundió algunos identificadores de partes XML (`sheetN.xml`) con los CodeName VBA. Son identificadores distintos: el nombre de la parte XML no determina el CodeName del módulo de hoja. La discrepancia queda resuelta documentalmente mediante el volcado estructural; no se renombra, mueve ni modifica ningún módulo VBA.

##### Resolución Factual de la Discrepancia de SOTO (Columnas 17 y 18)
La inspección directa del archivo XML de la hoja `RESIDENCIA SOTO` (`sheet23.xml`, celda Q1 y R1) y del código de persistencia en `modDatabase.bas:1938-1948` resuelve definitivamente la contradicción de versiones anteriores:
1. **Encabezados reales en la cuadrícula:** La Columna 17 (Q) se titula literalmente **`PRECIO APTO.`** y la Columna 18 (R) se titula **`SUPLE. OCUPAN.`**. No corresponden funcionalmente a habitaciones individuales ni dobles (conceptos hoteleros ajenos a los apartamentos de Soto del Barco), ni a observaciones.
2. **Tratamiento en persistencia:** El procedimiento `SincronizarFilaHaciaAccess` lee la fila completa hasta la columna 28 (`rowVals = ws.Range(ws.Cells(fila, 1), ws.Cells(fila, 28)).Value`), pero en el bloque condicional `ElseIf resTarget = "SOTO"` **omite deliberadamente asignar las posiciones `rowVals(1, 17)` y `rowVals(1, 18)` a ningún campo de Access**. Por tanto, dichos valores no se vuelcan a la tabla `Ordenes`.
3. **Columna 19 (S):** Su encabezado real es **`APARTAMENTO`** y se asigna al campo relacional `HabitacionesAsignadas` en Access (`AsignarCampoTexto rs, "HabitacionesAsignadas", rowVals(1, 19)`).

##### Diferenciación de Magnitudes: Columnas Físicas vs Campos Relacionales
- **Columnas de negocio:** 28 columnas (A:AB) en GIJÓN y SOTO; 28 columnas principales (A:AB) en OVIEDO, más Col. 32 (AF, notas locales).
- **Columnas técnicas y límites de navegación (`ScrollArea`):**
  - **GIJÓN:** Cols 29 (AC) y 30 (AD). El inventario previo atribuye a `Hoja2.cls:54` el límite de navegación **`A1:AD4000`** (Columna 30); CodeName confirmado: `Hoja2`.
  - **SOTO:** Cols 29 (AC) y 30 (AD). El inventario previo atribuye a `Hoja20.cls:57` el límite de navegación **`A1:AD4000`** (Columna 30); CodeName confirmado: `Hoja20`.
  - **OVIEDO:** Cols 29 (AC), 30 (AD), 31 (AE), 32 (AF) y 33 (AG). El inventario previo atribuye a `Hoja24.cls:29` el límite **`A1:AG4000`** (Columna 33); CodeName confirmado: `Hoja24`.
- **Columnas leídas en memoria por `SincronizarFilaHaciaAccess`:** Estrictamente **28 columnas (1 a 28, A:AB)** en las tres sedes. En OVIEDO, las columnas 29 a 33 (AC:AG) **no son leídas** por esta rutina.
- **Campos escritos en Access mediante `UPDATE`:**
  - **GIJÓN: 28 campos** (26 de datos de negocio + `FechaModificacion` y `UsuarioModificacion`). Col. 28 (`textoAB`) no se mapea.
  - **SOTO: 26 campos** (Cols 17, 18 y 28 leídas en memoria pero no asignadas a la tabla `Ordenes`).
  - **OVIEDO: 29 campos** (incluye `CamaSuple`, y Col. 28 mapea a `EstadoPago`).
- **Campos asignados en alta nueva (`AddNew`):** Se añaden 6 campos de clave y auditoría (`NumOrden`, `Residencia`, `FechaCreacion`, `UsuarioCreacion`, `Solapamiento`, `ConsentimientoRGPD`), totalizando 34 campos en GIJÓN, 32 campos en SOTO y 35 campos en OVIEDO.

#### 2.5. Autenticación, Permisos, Token de Instancia (CoCreateGuid) y Roles (ADMIN y OPERADOR)
- **Almacenamiento de Credenciales:**  
  La tabla `Usuarios` almacena las credenciales mediante el campo `Clave` como un **hash criptográfico SHA-256 irreversible de 64 caracteres hexadecimales** (`modDatabase.HashSHA256`). No se utiliza cifrado reversible ni almacenamiento en texto plano.
- **Roles Reales Implementados:**  
  La auditoría directa de la tabla `Usuarios` y del código VBA acredita la existencia de **únicamente dos roles reconocidos**:
  1. **`"ADMIN"`:** Rol de administración con privilegios totales (gestión de usuarios, tarifas, logs y purgas RGPD). La validación en el código fuente evalúa:
     ```vb
     If UCase(Trim(m_Rol)) = "ADMIN" Then ...
     ```
  2. **`"OPERADOR"`:** Rol operativo destinado a la captura de reservas, modificación ordinaria de expedientes y facturación en las sedes asignadas.
  - *Aclaración:* No existe en la base de datos ni en el código ningún rol denominado `"SoloLectura"`, ni se utiliza el literal `"Administrador"`.
- **Token de Sesión e Identidad de Instancia (`CoCreateGuid`):**  
  Implementado en `modDatabase.bas` (procedimiento `ObtenerTokenSesionInstancia`, líneas 1029-1059). El token de instancia de Excel (destinado a diferenciar operativamente instancias en condiciones normales de ejecución) se genera al iniciar la sesión invocando la API de Windows `CoCreateGuid` (`ole32.dll`):
  ```
  Formato general: usr & "@" & m_SessionTokenCache   (donde m_SessionTokenCache = pc & "#" & UCase$(hexGuid))
  ```
  - **Componentes, Aritmética de Longitudes y Limitaciones Demostradas:**
    1. *Extracción Parcial de 48 bits:* El código invoca `CoCreateGuid`, pero extrae exclusivamente los campos `Data1` (8 caracteres hex) y `Data2` (4 caracteres hex), obteniendo una cadena de 12 caracteres hexadecimales (48 bits). Los campos `Data3` y `Data4` de la estructura `GUID_TYPE` **no se utilizan**, por lo que **no constituye un identificador GUID completo de 128 bits**.
    2. *Mecanismo de Fallback Pseudo-Aleatorio Débil:* Si la llamada a la API `CoCreateGuid` falla (retorno distinto de 0), el código ejecuta:
       ```vb
       Randomize Timer
       hexGuid = Right$("00000000" & Hex$(CLng(Int((2147483647# * Rnd) + 1))), 8) & _
                 Right$("0000" & Hex$(CLng(Timer * 100) And &HFFFF&), 4)
       ```
       Este mecanismo de reserva recurre al generador pseudo-aleatorio `Rnd` inicializado con `Timer`, el cual **carece de garantías criptográficas de impredecibilidad y de unicidad absoluta**, pudiendo sufrir colisiones si dos puestos experimentan un fallo simultáneo de API.
    3. *Finalidad Operativa:* Su diseño responde a la necesidad de **diferenciar operativamente instancias de Excel en condiciones normales de ejecución**, pero no debe calificarse en ningún caso como un identificador único garantizado, criptográficamente inmune a colisiones.
    3. `pc` (1 a 12 caracteres): Obtenido de `Trim$(Environ$("COMPUTERNAME"))` (fallback `"PC"`). Si `Len(pc) > 12`, se recorta con `Left$(pc, 12)`.
    4. `m_SessionTokenCache` (14 a 25 caracteres): Concatena `pc & "#" & UCase$(hexGuid)`. Su longitud exacta es $\text{Len}(pc) + 1 + 12 = \text{Len}(pc) + 13$.
    5. `@`: Separador literal (1 carácter).
    6. `usr` ($\ge 1$ carácter): Obtenido de `Trim$(ObtenerUsuarioActual())` (fallback `Environ$("USERNAME")`, y si vacío `"admin"`).
  - **Fórmula de Longitud Sin Truncar:**
    $$\text{Longitud Sin Truncar} = \text{Len}(usr) + \text{Len}(pc) + 14$$
    (Para un equipo con nombre de 12 caracteres: $\text{Len}(usr) + 26$).
  - **Regla Estricta de Truncamiento y Límite Físico (50 Caracteres):**
    ```vb
    If Len(tokenCompleto) > 50 Then
        tokenCompleto = Left$(usr, 50 - Len(m_SessionTokenCache) - 1) & "@" & m_SessionTokenCache
    End If
    ```
    Cuando la longitud inicial supera 50 caracteres, el componente `usr` se recorta deterministamente a $50 - \text{Len}(m\_SessionTokenCache) - 1$ caracteres. La longitud total resultante es:
    $$\text{Longitud Truncada} = (50 - \text{Len}(m\_SessionTokenCache) - 1) + 1 + \text{Len}(m\_SessionTokenCache) = 50\text{ caracteres exactos}.$$
    El máximo matemático absoluto es **estrictamente 50 caracteres**, coincidiendo exactamente con la capacidad del campo `UsuarioModificacion Text(50)` en Access.
  - **Cálculo Exacto del Ejemplo Canónico:**
    Para el valor `jgarcia_operador@PCDESPACHO01#4A8F12B003C5`:
    - `usr` (`"jgarcia_operador"`): 16 caracteres.
    - `@`: 1 carácter.
    - `pc` (`"PCDESPACHO01"`): 12 caracteres.
    - `#`: 1 carácter.
    - `hexGuid` (`"4A8F12B003C5"`): 12 caracteres.
    - Longitud total exacta: $16 + 1 + 12 + 1 + 12 = \mathbf{42\text{ caracteres}}$ (se rectifica la estimación imprecisa de "46-47" de versiones anteriores).  
    Si `usr` tuviera el límite estándar de Active Directory de 20 caracteres (`SAMAccountName`), la longitud sería de **46 caracteres** ($20 + 26$); si tuviera 21 caracteres, sería de **47 caracteres**; y para cualquier longitud de usuario superior a 24 caracteres con un PC de 12 caracteres, se trunca de forma fija a **50 caracteres**.
  - **Distinción Conceptual Crítica:**
    1. **Identidad de Usuario (`m_NombreUsuario`):** Nombre del operador humano autenticado.
    2. **Token de Instancia (`usuario@PC#HexGUID`):** Identificador volátil de la instancia Excel en ejecución, utilizado como firma en `UsuarioModificacion` para el arrendamiento Lease Lock en Access.
    3. **CapKey V5.2 (`m_EngineCapKey`):** Clave criptográfica privada de 256 bits (64 hex) generada mediante CSPRNG y utilizada exclusivamente en laboratorio como oráculo interno de ejecución transaccional.

#### 2.6. Persistencia y Operaciones ADODB Directas (32 Procedimientos en modDatabase)
La interacción entre el libro Excel y la base de datos Access en producción se ejecuta mediante **32 procedimientos de modificación directa** ubicados en `modDatabase.bas`.
- Cada procedimiento gestiona de forma autónoma conexiones ADODB puntuales utilizando cadenas de conexión con el proveedor `Microsoft.ACE.OLEDB.12.0`.
- Determinados procedimientos implementan transacciones locales mediante `cn.BeginTrans`, `cn.CommitTrans` y `cn.RollbackTrans` (ej. `ReservarSiguienteOrdenLibreBD`, `SincronizarFilaHaciaAccess`, `EliminarOrdenBD`), mientras que otros ejecutan sentencias DML directas sin bloque transaccional explícito.
- El inventario exhaustivo línea a línea de estos 32 procedimientos, sus tablas afectadas, validaciones y llamadores se encuentra documentado íntegramente en el **[ANEXO A](#anexo-a--inventario-exhaustivo-de-operaciones-adodb-de-producción-32-procedimientos)** de este documento.

#### 2.7. Gestión Operativa del Nº ORDEN en Producción
En el código de producción actual, la asignación de números de orden oficiales y provisionales se rige por:
- **Reserva Online Conectada (`ReservarSiguienteOrdenLibreBD`, `modDatabase.bas:1403-1472`):**  
  Ejecuta la consulta SQL:
  ```sql
  SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '...' AND NumOrden < 900001
  ```
  y propone `candidato = MAX + 1` en un bucle limitado a cinco intentos (`maxIntentos = 5`, `modDatabase.bas:1425-1429`) para intentar una inserción transaccional en `Ordenes`. En cada intento vuelve a calcular `MAX + 1`; la rutina no busca huecos interiores.
- **Comportamiento As-Built Actual vs Arquitectura Objetivo en el Caso Límite `900000`:**
  1. *Comportamiento As-Built en Producción:*
     - En `ReservarSiguienteOrdenLibreBD` (`modDatabase.bas:1403-1472`), si el máximo existente en Access es `899999`, al consultar `WHERE NumOrden < 900001`, la función calcula `candidato = 899999 + 1 = 900000` e **intenta insertar y consolidar el número 900000 como orden oficial en Access**, al carecer de corte en 899999.
     - En entrada manual conectada (`GestionarEntradaNumOrdenHoja`, `modDatabase.bas:1527-1696`), `EsNumOrdenProvisional(900000)` evalúa `(900000 >= 900001)` = **`False`**. En consecuencia, el sistema en línea lo procesa como número oficial ordinario e invoca `IntentarReservarOrdenBD`, consolidándolo si está libre en Access.
     - En entrada manual desconectada (modo offline), al no ser evaluado como provisional, el sistema activa la rama de contingencia y genera un número provisional nuevo mediante `ObtenerSiguienteNumOrdenProvisional(ws)` (e.g. 900001).
     - La función `EsNumOrdenProvisional` carece de cota superior, evaluando como provisional cualquier número $\ge 900001$ (incluso valores $\ge 1000000$).
  2. *Delimitación Canónica de la Arquitectura Objetivo (Aprobada en Diseño):*
     - **Oficiales Ordinarios:** Intervalo entero **`[1 .. 899.999]`**.
     - **Separador de Seguridad Excluido:** Valor **`900.000`** (estrictamente prohibido tanto en generación automática como en entrada manual online/offline).
     - **Provisionales de Contingencia (SICO-N2):** Intervalo entero acotado **`[900.001 .. 999.999]`**.
     - **Inválidos / Fuera de Rango:** Valores $\le 0$ o **$\ge 1.000.000$** (rechazo formal previo para evitar errores de desbordamiento `Error 6 Overflow` en conversiones a tipo `Long`).
- **Asignación Provisional Offline (`ObtenerSiguienteNumOrdenProvisional`, `ModuloSeguridadLocal.bas:598-615`):**  
  Recibe como parámetro el objeto `ByVal wsResidencia As Worksheet` y examina directamente la columna A de la hoja de cálculo de la residencia en memoria (de la fila 2 a la última ocupada) para identificar el valor máximo superior o igual a `MIN_ORDEN_PROVISIONAL` (900001) y retorna `maxProv + 1` (no consulta la cola WAL en disco).

#### 2.8. Gestión de Reservas, Huéspedes y Facturación
- **Hojas Maestras de Órdenes:** Las órdenes de alojamiento se registran en las hojas `GIJÓN`, `SOTO` y `OVIEDO`.
- **Validaciones en Entrada:** El evento `Worksheet_Change` supervisa las celdas críticas. Si se introduce un DNI o un Nombre que coincide con la tabla `ListaNegra`, el sistema interrumpe el registro y muestra una alerta modal.
- **Módulo de Facturación:** La función `ObtenerSiguienteNumFactura` (`modDatabase.bas:2678-2778`) obtiene y actualiza el correlativo anual de factura de forma concurrente sobre la tabla `ContadorFacturas`.

#### 2.9. Gestión de Tarifas, Precios y Turnos (Hoja28 / Turnos-Precios y TARIFAS UNIDAD)
- **Correspondencia Canónica de Hojas:**
  - **Hoja `Turnos-Precios`:** CodeName interno en el proyecto VBA: **`Hoja28`**. Nombre visible en la pestaña de Excel: **`Turnos-Precios`**. Contiene la configuración estacional de precios y fechas de turnos para la Residencia Soto del Barco.
  - **Hoja `TARIFAS UNIDAD`:** Contiene los precios desglosados por tipo de habitación para GIJÓN y OVIEDO, junto al cuadro de IVA.
- **Persistencia Base Existente en Access (`ConfigTarifasCeldas`):**  
  El sistema cuenta con persistencia celda a celda en la tabla `ConfigTarifasCeldas` (10 campos relacionales). La consulta de solo lectura realizada sobre producción en la Orden 1K (16/11/2026) obtuvo **291 registros activos**, distribuidos en `TARIFAS UNIDAD` (155) y `TURNOS-PRECIOS` (136).
  - **Rangos efectivos de `TARIFAS UNIDAD`:** `A3:C7` (15 posiciones), `E3:N9` (70) y `E13:N19` (70). No se detectaron fórmulas en esas posiciones; las 155 se encuentran representadas en Access.
  - **Rangos efectivos de `Turnos-Precios`:** `A1:F20` (120 posiciones, de las que 17 contienen fórmulas y no se persisten, quedando 103 registros), `A22:F23` (12) y `B25:D31` (21). Total: 136 registros.
  - **Conciliación de coordenadas:** las 42 posiciones que una comprobación anterior había clasificado como externas a los rangos de referencia truncados (`E3:J9` y `E13:L19`) pertenecen en realidad a las columnas `K:N` de los bloques GIJÓN/OVIEDO, incluidas en los rangos efectivos del código (`E3:N9` y `E13:N19`). La consulta 1K no identificó registros fuera de los rangos efectivos.
  *Distinción técnica fundamental:* Debe diferenciarse la persistencia base de tarifas ya operativa en Access de la propuesta de optimización de concurrencia granular **TAR-10**, la cual continúa en estado de diseño y pendiente de inyección.
- **Celdas Combinadas en Turnos-Precios:**  
  El procedimiento `CargarTarifasDesdeBD` en `modDatabase.bas` incorpora validación expresa contra errores de manipulación en el rango combinado `B25:D25`.

#### 2.10. Calendarios, Habitaciones y Bloqueos
- **Visualización Gráfica:** Matrices de ocupación mensual (habitaciones en filas vs. días en columnas), renderizadas mediante formatos condicionales y rutinas de dibujo VBA.
- **Ocultación y Visibilidad:** La tabla `EstadoHabitacionesCalendario` almacena las filas visibles y ocultas por residencia.
- **Bloqueos Temporales:** La tabla `BloqueosHabitaciones` registra los periodos de indisponibilidad por mantenimiento, obras o averías.

#### 2.11. Subsistema de Contingencia Desconectada (SICO-N2) y Registro Local WAL
- **Objetivo Funcional:** Proporcionar continuidad operativa permitiendo registrar y modificar expedientes de forma local cuando la red local (SMB) o el archivo Access no están accesibles.
- **Mecanismo de Registro Local:**  
  Las operaciones efectuadas en contingencia se almacenan en la hoja de cálculo oculta `SYS_COLA_SYNC` y se persisten en disco como archivos de texto plano con extensión `.wal` en el directorio local:
  `%LOCALAPPDATA%\ResidenciaApp\SICO-N2\`
- **Naturaleza del Archivo WAL:**  
  El archivo WAL de SICO-N2 es un **registro de texto plano estructurado en 13 columnas delimitadas**, gestionado por código VBA a nivel de aplicación para almacenamiento intermedio local. No proporciona garantías físicas de durabilidad ACID como las de un motor transaccional de base de datos.
- **Tratamiento del Valor `900000` (As-Built vs Objetivo):**  
  En el código As-Built actual de producción, el valor `900000` **no se encuentra físicamente bloqueado**: puede ser emitido por `ReservarSiguienteOrdenLibreBD` si el máximo es 899999 y es aceptado en línea por `IntentarReservarOrdenBD`. La regla de exclusión estricta de `900000` como separador de seguridad inhabilitado constituye una **especificación de la Arquitectura Objetivo** pendiente de inyección formal.

#### 2.12. Ciclo de Vida de Datos, Archivado y Purgas (5 Procedimientos Reales)
La auditoría exhaustiva del código fuente en `modDatabase.bas` y del esquema físico en Access acredita que el ciclo de vida de datos se articula en **cinco procedimientos específicos** con salvaguardas transaccionales y de negocio verificadas.

> [!IMPORTANT]
> **Esquema Real de `Ordenes_Historico` y Naturaleza del Borrado Previo:**  
> Como acredita la inspección ADOX/OLEDB (Apartado 6), la tabla `Ordenes_Historico` **carece de clave primaria** y su único índice `idx_Hist_Res_NumOrden` sobre `(Residencia, NumOrden)` está configurado con **`Unique = False`**. Por tanto, el borrado previo de registros históricos que ejecuta el código no responde a ninguna restricción técnica de Access para evitar un Error 3022 o colisión de clave, sino a una **política procedimental de sustitución o depuración de registros coincidentes** definida en la lógica de aplicación.

A continuación se detalla el comportamiento literal de los tres procedimientos que ejecutan este borrado previo:

1. **`EliminarOrdenBD` (`modDatabase.bas:4261-4420`):**  
   - *Adquisición de Lease Lock:* `UPDATE` condicional previo sobre `Ordenes` con el token de sesión. Si está en edición activa por otro operador (`recsLockedDel = 0`), cancela el borrado con alerta modal y segundos restantes.
   - *Condición de Datos Reales de Huésped:* Comprueba `Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0`. Solo si contiene datos reales se procede al archivado; las pre-reservas en blanco canceladas no se archivan en el histórico.
   - *Registros Históricos Eliminados antes de Insertar:* Ejecuta:
     ```sql
     DELETE FROM Ordenes_Historico WHERE NumOrden = N AND Residencia = '...'
     ```
     Esta sentencia elimina físicamente cualquier registro previo en `Ordenes_Historico` con esa combinación exacta. **Consecuencia destructiva:** si el número de orden $N$ fue reutilizado en la cuadrícula activa (p. ej. tras una baja previa o por saltos en la numeración), este borrado elimina y destruye de forma irreversible el expediente histórico legítimo anterior que portaba ese mismo número, sustituyéndolo por el nuevo.
   - *Registros Activos Copiados:* Copia a `Ordenes_Historico` los datos del expediente activo de la cuadrícula mediante `INSERT INTO Ordenes_Historico (...) VALUES (...)`.
   - *Registros Activos Eliminados Posteriormente:* Ejecuta `DELETE FROM Ordenes WHERE NumOrden = N AND Residencia = '...' AND [UsuarioModificacion] = 'token'` comprobando `recsDeleted > 0`.
   - *Ámbito Transaccional:* Las tres operaciones DML están encapsuladas dentro de una **única transacción atómica** (`cn.BeginTrans` / `cn.CommitTrans`).

2. **`EliminarOrdenesLoteBD` (`modDatabase.bas:4424-4552`):**  
   - *Registros Históricos Eliminados:* En bloques de hasta 200 órdenes, ejecuta `DELETE FROM Ordenes_Historico WHERE Residencia = '...' AND NumOrden IN (N1, N2, ...)`.
   - *Registros Activos Copiados:* Inserta en `Ordenes_Historico` las órdenes del lote que contienen datos reales de huésped (`Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0`).
   - *Registros Activos Eliminados:* Ejecuta `DELETE FROM Ordenes WHERE Residencia = '...' AND NumOrden IN (N1, N2, ...)`.
   - *Ámbito Transaccional:* Cada lote de 200 expedientes se procesa bajo `cn.BeginTrans` / `cn.CommitTrans`.
   - *Consecuencia:* Comparte idéntica política de sustitución: si algún número del lote coincide con un registro histórico anterior, dicho registro histórico se destruye físicamente.

3. **`EliminarTodasOrdenesSedeBD` (`modDatabase.bas:4556-4659`):**  
   Vaciado administrativo exclusivo de los expedientes activos de una residencia (`WHERE Residencia = '...'`). **Implementa fáctica y estrictamente el Caso B:**
   - *Comprobación previa:* Verifica que ningún expediente activo de la sede tenga Lease Lock activo ($< 120$ s) de otra sesión.
   - *Sentencia 1 (Sustitución de Coincidentes):* Ejecuta:
     ```sql
     DELETE FROM Ordenes_Historico WHERE Residencia = '...' 
     AND NumOrden IN (SELECT NumOrden FROM Ordenes WHERE Residencia = '...' 
                      AND (Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0))
     ```
     Elimina **exclusivamente los registros históricos cuyos números de orden coincidan con las órdenes activas con datos reales que van a ser archivadas** en la misma transacción. **No elimina el histórico consolidado previo de la sede no coincidente**, el cual permanece íntegro.
   - *Sentencia 2 (Archivado en Bloque):* Ejecuta `INSERT INTO Ordenes_Historico SELECT ... FROM Ordenes WHERE Residencia = '...' AND (Len(Trim(DNI)) > 0 Or Len(Trim(Nombre)) > 0)`.
   - *Sentencia 3 (Vaciado de Activas):* Ejecuta `DELETE FROM Ordenes WHERE Residencia = '...'`.
   - *Ámbito Transaccional:* Todo el proceso está encapsulado bajo `cn.BeginTrans` / `cn.CommitTrans`.
4. **`ArchivarOrdenesAntiguasBD` (`modDatabase.bas:4784-4970`):**  
   Procedimiento orquestador compuesto que opera en **dos fases formalmente desacopladas**:
   - *Fase 1 (Persistencia en Access):* Identifica expedientes con antigüedad superior a `diasRetencion` (default 90 días) y delega el DML transaccional en `EliminarOrdenesLoteBD` (Entrada #21 del inventario) para archivar en `Ordenes_Historico` y borrar de `Ordenes`.
   - *Fase 2 (Compactación Visual en Excel):* Desactiva eventos, refresco y cálculo (`ScreenUpdating = False`, `EnableEvents = False`, `Calculation = xlCalculationManual`), suspende transitoriamente el cálculo en hojas satélite dependientes (`wsCal.EnableCalculation = False`, `wsSum.EnableCalculation = False`), desprotege la hoja, agrupa filas en bloques de hasta 100 mediante `Application.Union` y las elimina con `Delete Shift:=xlUp` (o `ClearContents` preservando fórmulas si se archiva la hoja completa). Restaura la protección con `UserInterfaceOnly:=True`, reactiva cálculos satélite de forma escalonada con `DoEvents` y restituye el cálculo y refresco global, registrando la operación en `LogActividad`.
5. **`PurgarHistoricoRGPD` (`modDatabase.bas:4973-5003`):**  
   Procedimiento de purga física invocable a demanda (desde `frmPanelRGPD`), que elimina de forma definitiva registros históricos cuya fecha de salida supere el parámetro numérico de retención (default en signatura `anosRetencion = 3`):
   ```vb
   sqlDelete = "DELETE FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
   ```
   (purgando también `BDAS_Historico_Snapshot` si estuviera presente) y emite un registro inmutable en `LogActividad`. El fundamento y adecuación jurídica de este plazo queda formalmente deslindado como una decisión organizativa y legal externa.

#### 2.13. Limitaciones Funcionales y Defectos Estructurales de Producción
1. **Condiciones de Carrera en Numeración:** Asignación de números mediante `MAX+1` sin búsqueda atómica de huecos interiores.
2. **Vulnerabilidad en Reconciliación FIFO:** El procedimiento `ResolverNumOrdenProvisionalEnBD` puede generar colisiones al reconciliar números de contingencia.
3. **Persistencia Directa Descentralizada:** 32 procedimientos independientes atacando Access directamente sin una pasarela transaccional unificada.
4. **Dependencias No Inyectadas:** La clase transaccional `clsDatabaseEngine` (Fase 4) permanece en archivos de laboratorio sin haber sido consolidada en el binario maestro de producción.
5. **Concurrencia Parcial en Importación Masiva:** `SincronizarLoteImportacionBD` opera dentro de una transacción, pero no valida ni adquiere el Lease Lock individual antes de actualizar registros existentes (`DEF-FIFO-07`).

---

### 3. BLOQUE II — ARQUITECTURA OBJETIVO APROBADA

#### 3.1. Marco Arquitectónico V5.2 y Principios Rectores
El Marco Arquitectónico V5.2 constituye la referencia técnica aprobada para la evolución estructural de ResidenciaApp. Su propósito es subsanar de raíz las deficiencias de concurrencia, sobreescrituras accidentales y dispersión de persistencia detectadas en la versión As-Built.
- **Principio de Mediación Única:** Ningún componente de interfaz (hoja de cálculo o formulario) debe emitir sentencias SQL directas contra el archivo Access. Toda operación de persistencia debe canalizarse a través de un motor transaccional especializado.
- **Delimitación Transaccional y Coordinación Procedimental:**  
  La arquitectura reconoce formalmente que las transacciones ACID estrictas (`BeginTrans`, `CommitTrans`, `RollbackTrans`) operan exclusivamente en el motor Access a través del proveedor OLEDB. Las modificaciones sobre las celdas de las hojas de Excel se gestionan mediante **orquestación procedimental acoplada**:
  1. Validación rigurosa de tipos y precondiciones en memoria.
  2. Ejecución de la transacción SQL en Access.
  3. Si la transacción en Access se confirma (`Commit`), se actualizan los valores en las celdas de la cuadrícula de Excel.
  4. Si la transacción falla o detecta colisión, se aborta (`Rollback`), se cancela la mutación en Excel y se restaura el estado visual original mediante compensación.
- **Cero Pérdida de Datos y Trazabilidad:** Cada modificación debe dejar constancia del token de sesión de la instancia y de la marca temporal exacta de la transacción.

#### 3.2. Motor Transaccional Especializado: `clsDatabaseEngine` (Estado Real de Laboratorio)
- **Naturaleza, Presencia en Laboratorio y Estado de Homologación:**  
  Clase VBA (`clsDatabaseEngine.cls`) presente físicamente en el proyecto VBA del binario de laboratorio `BDAS_Multiusuario_Fase4.xlsm` (SHA-256: `299b50fb...`, sumando 175 componentes).  
  *Deslinde de Certeza y Validación Funcional:* Aunque el código de la clase (490 líneas) ha sido extraído e inspeccionado en laboratorio, **su homologación funcional completa no se encuentra acreditada con registros detallados y reproducibles de prueba en la documentación permanente del proyecto**. Por tanto, la presencia física de la clase en el binario de pruebas es un hecho acreditado, pero su estado de validación operativa permanece clasificado bajo **Certeza 4 (Incertidumbre / Pendiente de Evidencia de Pruebas)** según la taxonomía del apartado 8.1, sin que ello implique que la clase sea errónea o inoperativa.
- **Ciclo de Vida de Conexiones (Apertura y Cierre Puntual por Operación):**  
  La inspección directa del código fuente de `clsDatabaseEngine.cls` acredita que el motor **no mantiene una conexión ADODB persistente** durante la vida de la instancia. En su lugar, **abre y cierra conexiones ADODB puntuales por cada operación ejecutada**:
  - En `ExecuteNonQuery` (líneas 233-239): ejecuta `Set cn = CreateObject("ADODB.Connection")`, `cn.Open modDatabase.GetConnectionString()`, ejecuta el comando SQL y de inmediato ejecuta `cn.Close` y libera el objeto.
  - En `EjecutarUpdateCondicionalOrden` (líneas 411-450): crea la conexión, la abre con timeout de 5 segundos, ejecuta la sentencia `cn.Execute sql, affected` e inmediatamente ejecuta `cn.Close` y `Set cn = Nothing`.
  Por tanto, se debe distinguir con precisión la **instancia singleton de la clase** (`m_DbEngine` en `modDatabase.bas`), la cual permanece instanciada durante la sesión, de las **conexiones ADODB**, las cuales son efímeras y se crean, abren y destruyen en cada llamada a método.
- **Delimitación Estricta y Firmas Completas sin Elipsis:**  
  A continuación se documentan las firmas literales completas separando rigurosamente los componentes de la clase encapsulada de los de la fachada:

  **1. Miembros Físicamente Implementados en `clsDatabaseEngine.cls`:**
  | Procedimiento / Función | Ámbito | Firma Completa Exacta | Retorno | Responsabilidad Exclusiva |
  | :--- | :--- | :--- | :---: | :--- |
  | `ValidarCapacidadInterna` | `Private` | `Private Function ValidarCapacidadInterna(ByVal capKey As String) As Boolean` | `Boolean` | Valida en tiempo constante si la `capKey` presentada coincide con la custodiada por la fachada mediante `modDatabase.VerificarCapacidadMotorInterno`. |
  | `VincularInstanciaUnica` | `Friend` | `Friend Function VincularInstanciaUnica(ByVal instanceId As String, ByVal capKey As String) As Boolean` | `Boolean` | Vincula de forma indivisible el motor a la instancia autorizada. Impide re-vinculaciones posteriores. |
  | `ExecuteNonQuery` | `Friend` | `Friend Function ExecuteNonQuery(ByVal sql As String, Optional ByVal capKey As String = vbNullString) As Boolean` | `Boolean` | Ejecuta comandos DML en Access bajo validación previa de CapKey. **Retorna Boolean**, no `Long`. |
  | `EjecutarUpdateCondicionalOrden` | `Friend` | `Friend Function EjecutarUpdateCondicionalOrden(ByVal resTarget As String, ByVal numOrden As Long, ByVal setParts As String, Optional ByRef outUsrBloqueo As String = "", Optional ByRef outSegRestantes As Long = 0, Optional ByRef outExisteExpediente As Boolean = True, Optional ByVal capKey As String = vbNullString) As Boolean` | `Boolean` | Ejecuta actualizaciones optimistas verificando condiciones de Lease Lock V5.2 y firma de sesión, devolviendo por referencia los datos del bloqueo concurrente. |

  **2. Miembros Implementados en la Fachada `modDatabase.bas` (Laboratorio Fase 4, Líneas 18045-18307):**
  | Miembro / Procedimiento | Ámbito | Declaración / Firma Completa Exacta | Responsabilidad Exclusiva |
  | :--- | :--- | :--- | :--- |
  | `m_DbEngine` | `Private` | `Private m_DbEngine As clsDatabaseEngine` (Línea 29) | Instancia singleton del motor transaccional. |
  | `m_EngineInstanceId` | `Private` | `Private m_EngineInstanceId As String` (Línea 31) | Identificador único de instancia (32 caracteres hexadecimales). |
  | `m_EngineCapKey` | `Private` | `Private m_EngineCapKey As String` (Línea 33) | Clave de capacidad CSPRNG de 256 bits (64 caracteres hexadecimales). |
  | `m_EngineCapVerifier` | `Private` | `Private m_EngineCapVerifier As String` (Línea 35) | Verificador SHA-256 de la clave de capacidad (`HashSHA256(capKey)`). |
  | `CompararCadenasTiempoConstante` | `Public` | `Public Function CompararCadenasTiempoConstante(ByVal a As String, ByVal b As String) As Boolean` (Líneas 18045-18067) | Comparación segura mediante acumulación XOR de diferencias de longitud y caracteres UTF-16, sin cortocircuito para mitigar ataques de temporización. |
  | `GenerarCredencialCSPRNG64Hex` | `Private` | `Private Function GenerarCredencialCSPRNG64Hex() As String` (Líneas 18071-18121) | Generador criptográficamente seguro de 32 bytes (256 bits) utilizando la API `BCryptGenRandom` con fallback a `RtlGenRandom` (`SystemFunction036`). Sin fallback débil; si ambos fallan, eleva error 5101. |
  | `GenerarGuidHex32` | `Private` | `Private Function GenerarGuidHex32() As String` (Líneas 18125-18165) | Generador de identificador único de instancia invocando `CoCreateGuid` y formateando los 16 bytes a 32 caracteres hexadecimales en minúsculas. |
  | `RegistrarCapacidadMotorInterno` | `Private` | `Private Sub RegistrarCapacidadMotorInterno(ByVal capKey As String, ByVal verifier As String)` (Líneas 18169-18175) | Almacena privadamente en las variables de módulo la credencial y su verificador hash. |
  | `VerificarCapacidadMotorInterno` | `Public` | `Public Function VerificarCapacidadMotorInterno(ByVal instanceId As String, ByVal capKey As String) As Boolean` (Líneas 18179-18217) | Punto de validación cruzada invocado por `clsDatabaseEngine.ValidarCapacidadInterna`. Comprueba longitudes (32, 64, 64) y compara en tiempo constante instancia, clave y verificador hash. |
  | `ObtenerMotorBaseDatos` | `Private` | `Private Function ObtenerMotorBaseDatos() As clsDatabaseEngine` (Líneas 18221-18307) | Factoría singleton interna con ámbito **`Private`**. Verifica la integridad de la instancia existente; si no existe o detecta corrupción, instancia `New clsDatabaseEngine`, genera credenciales por CSPRNG y ejecuta `VincularInstanciaUnica`. |
  | `EjecutarDMLInterno` | `Private` | `Private Function EjecutarDMLInterno(ByVal sql As String) As Boolean` (Líneas 4817-4855) | Procedimiento consumidor que obtiene el motor mediante `ObtenerMotorBaseDatos()` y suministra `m_EngineCapKey` a `engine.ExecuteNonQuery(sql, m_EngineCapKey)`. |
  | `EjecutarUpdateCondicionalOrden` | `Private` | `Private Function EjecutarUpdateCondicionalOrden(ByVal resTarget As String, ByVal numOrden As Long, ByVal setParts As String, Optional ByRef outUsrBloqueo As String = "", Optional ByRef outSegRestantes As Long = 0, Optional ByRef outExisteExpediente As Boolean = True) As Boolean` (Líneas 8355-8395) | Fachada que obtiene el motor mediante `ObtenerMotorBaseDatos()` y delega en `engine.EjecutarUpdateCondicionalOrden(...)` suministrando `m_EngineCapKey`. Su ámbito en `modDatabase.bas` es estrictamente `Private` (no `Public` ni `Friend`). |
- **Estado de Inyección:** **0% EN PRODUCCIÓN.** Esta clase permanece aislada en el entorno de pruebas Fase 4 y no ha sido inyectada aún en `BDAS_Multiusuario.xlsm`.

#### 3.3. Clave de Capacidad (CapKey de 256 bits), Ciclo de Vida y Escritura Protegida
- **Generación, Transmisión y Validación de la CapKey:**  
  1. *Generación Criptográfica:* En `modDatabase.ObtenerMotorBaseDatos()`, la fachada invoca `GenerarCredencialCSPRNG64Hex()`, la cual utiliza llamadas CSPRNG nativas de Windows (`BCryptGenRandom` y `RtlGenRandom`) para generar 32 bytes de entropía pura expresados como 64 caracteres hexadecimales en minúsculas. Si ambas APIs del sistema operativo fallan, eleva un error de ejecución sin recurrir a funciones pseudo-aleatorias débiles.
  2. *Transmisión Inicial:* La fachada instancia `newEngine` y le transmite la clave y el identificador de 32 hex (`GenerarGuidHex32()`) mediante `newEngine.VincularInstanciaUnica(instId, capKey)`. La vinculación es indivisible y de un solo uso.
  3. *Invocación Protegida:* Las operaciones modificadoras internas (`EjecutarDMLInterno` en línea 4835 y `EjecutarUpdateCondicionalOrden` en línea 8373) obtienen la instancia con `ObtenerMotorBaseDatos()` y transfieren privadamente `m_EngineCapKey` como argumento a `engine.ExecuteNonQuery` o `engine.EjecutarUpdateCondicionalOrden`.
  4. *Validación Oracular en Tiempo Constante:* Como primera línea de ejecución, el método del motor invoca `ValidarCapacidadInterna(capKey)`, la cual valida el hash local en tiempo constante y recurre a `modDatabase.VerificarCapacidadMotorInterno`, comprobando la clave mediante `CompararCadenasTiempoConstante(a, b)` con acumulación XOR para prevenir ataques por canal lateral (*timing attacks*). Si la clave no es válida, la ejecución se interrumpe de inmediato retornando `False`.
  *Deslinde de Login:* La CapKey es un token efímero de infraestructura de software generado en memoria durante el arranque del motor transaccional; no interviene en `frmLogin`, el cual se limita a cotejar el usuario y su hash SHA-256 en la tabla `Usuarios`.

#### 3.4. Control de Concurrencia Optimista y Lease Lock (120 Segundos)
Para resolver las colisiones entre operadores concurrentes que editan la misma reserva en sedes distintas o simultáneamente:
- **Arrendamiento de Edición (Lease Lock):**  
  Durante la ejecución de `SincronizarFilaHaciaAccess`, el sistema intenta adquirir un arrendamiento temporal de **120 segundos** sobre el registro correspondiente en Access, grabando en la tabla `Ordenes`:
  - `UsuarioModificacion`: Token de sesión de la instancia con formato `usuario@PC#HexGUID` (acotado a <= 50 caracteres).
  - `UltimaModificacion`: Marca temporal del sistema (`Now`).
- **Condición Atómica de Escritura (UPDATE Condicional):**  
  La persistencia final del expediente se ejecuta mediante una cláusula SQL condicional. En cumplimiento estricto de la Orden 1G, a continuación se deslindan con exactitud matemática las dos implementaciones reales existentes (laboratorio V5.2 vs producción), contrastando minuciosamente cada uno de sus predicados de búsqueda:

  **1. Sentencia SQL Real Extraída de `clsDatabaseEngine.cls` (Líneas 365-433 en Laboratorio Fase 4):**
  ```sql
  UPDATE Ordenes SET <setParts>, [UsuarioModificacion] = '<tokenMod>', [UltimaModificacion] = #yyyy-mm-dd hh:nn:ss#
  WHERE NumOrden = <numOrden> AND Residencia = '<resTarget>'
    AND ([UsuarioModificacion] IS NULL 
         OR [UsuarioModificacion] = '' 
         OR [UsuarioModificacion] = '<tokenMod>' 
         OR [UltimaModificacion] IS NULL 
         OR DateDiff('s', [UltimaModificacion], Now()) >= 120)
  ```
  *Condiciones evaluadas en laboratorio V5.2:*
  - Admite registros con `UsuarioModificacion IS NULL`.
  - Admite registros con `UsuarioModificacion = ''`.
  - Admite el propio token de la instancia actual (`tokenMod`).
  - Admite `UltimaModificacion IS NULL`.
  - Admite un arrendamiento cuya antigüedad sea igual o superior a 120 segundos (`DateDiff('s', [UltimaModificacion], Now()) >= 120`).
  *(Nota crítica: La implementación V5.2 de laboratorio **no incorpora** la condición de marca temporal futura `DateDiff < 0` ni contempla la variable de usuario de fallback en esta sentencia).*

  **2. Sentencia SQL Real en Código de Producción (`modDatabase.bas:1791-1800`):**
  En producción, `SincronizarFilaHaciaAccess` ejecuta previamente la adquisición del Lease Lock mediante:
  ```sql
  UPDATE Ordenes 
  SET [UsuarioModificacion] = '<tokenActual>', 
      [UltimaModificacion] = #yyyy-mm-dd hh:nn:ss# 
  WHERE NumOrden = <nOrden> AND Residencia = '<resTarget>' 
    AND ([UsuarioModificacion] IS NULL 
         OR [UsuarioModificacion] = '' 
         OR [UsuarioModificacion] = '<tokenActual>' 
         OR [UltimaModificacion] IS NULL 
         OR DateDiff('s', [UltimaModificacion], Now()) >= 120 
         OR DateDiff('s', [UltimaModificacion], Now()) < 0)
  ```
  *Condiciones evaluadas en producción:*
  - Usuario de modificación nulo (`[UsuarioModificacion] IS NULL`).
  - Usuario de modificación vacío (`[UsuarioModificacion] = ''`).
  - Token de la sesión de la instancia actual (`<tokenActual>`).
  - Marca temporal nula (`[UltimaModificacion] IS NULL`).
  - Antigüedad igual o superior a 120 segundos (`DateDiff('s', [UltimaModificacion], Now()) >= 120`).
  - **Marca temporal futura respecto al reloj del sistema que evalúa la condición (`DateDiff('s', [UltimaModificacion], Now()) < 0`):** condición exclusiva de producción destinada a absorber desfases horarios de reloj entre puestos clientes.

- **Comportamiento Funcional y Escenarios de Adquisición:**  
  La ejecución del `UPDATE` condicional debe interpretarse distinguiendo cinco situaciones operativas:

| Escenario | Comportamiento que debe documentarse |
|---|---|
| **1. Registro existente con Lease Lock vigente de otra instancia** | Solo se garantiza `affected = 0` cuando `UsuarioModificacion` contiene un titular no nulo y no vacío, distinto de `tokenActual`; `UltimaModificacion` no es nulo; y `DateDiff('s', UltimaModificacion, Now())` está entre **0 y 119**, inclusive. El SQL de producción no compara `UsuarioModificacion` con el nombre de usuario (`uActual`). Bajo esas precondiciones ninguna alternativa del `WHERE` permite adquirir el registro y el `UPDATE` afecta a **cero filas**. |
| **2. Registro existente y Lease Lock caducado (120 segundos o más)** | Si el arrendamiento temporal ha expirado (`DateDiff('s', UltimaModificacion, Now()) >= 120`), la condición temporal se evalúa como verdadera. Si existe una única fila para la clave compuesta `(NumOrden, Residencia)` y la ejecución termina correctamente, la sentencia permite adquirir el bloqueo y devuelve `recsLock = 1`, transfiriendo la titularidad a la instancia actual y renovando su token y marca temporal. Por tanto, es técnicamente incorrecto afirmar que un bloqueo caducado produce cero registros afectados: el bloqueo caducado habilita la toma de control del expediente. |
| **3. Registro existente y marca temporal futura** | La implementación de producción contempla expresamente `DateDiff('s', [UltimaModificacion], Now()) < 0` como condición adicional de adquisición, permitiendo tomar el control si el reloj cliente grabó una fecha adelantada; **esa condición no figura en el SQL V5.2 de laboratorio**. Una marca temporal futura no debe describirse en ningún caso como un bloqueo caducado, tratándose de una situación técnica diferenciada. |
| **4. Registro inexistente o clave de búsqueda no coincidente** | Si la clave `(NumOrden, Residencia)` no existe, el `UPDATE` afecta a cero filas (`recsLock = 0`). La rama `If recsLock <> 1` ejecuta un `SELECT` posterior, pero no diferencia el resultado `EOF` del caso de bloqueo: si no encuentra fila, conserva los valores iniciales (`usrLock = "otro usuario"` y el tiempo restante predeterminado), muestra igualmente el mensaje genérico de conflicto, revierte la transacción, llama a `SincronizarOrdenHaciaExcel(ws, numOrden, fila, resTarget)` y sale de la función. La firma de `SincronizarFilaHaciaAccess` no contiene un parámetro `permitirAlta`. Por tanto, el código no implementa en esta rama una recuperación `AddNew` para una clave inexistente. |
| **5. Metadatos que habilitan la adquisición sin esperar 120 segundos** | La sentencia permite adquirir si `UsuarioModificacion` es `NULL`, está vacío o coincide con `tokenActual`, o si `UltimaModificacion` es `NULL`. Estas alternativas prevalecen aunque la fecha no haya alcanzado 120 segundos; en particular, `UltimaModificacion IS NULL` permite la adquisición aunque exista otro titular no vacío. No deben presentarse como un bloqueo vigente protegido. |

**Tratamiento operativo después de `recsLock <> 1`:** la función consulta `UsuarioModificacion` y `UltimaModificacion`, pero no comprueba una rama específica para `rsLock.EOF`. En ambos casos (titular existente o clave inexistente) muestra el mismo aviso genérico de conflicto, revierte la transacción, refresca la fila mediante `SincronizarOrdenHaciaExcel` y retorna `False`. La llamada de refresco y el mensaje no permiten concluir que el código haya identificado correctamente al titular cuando la fila no existe.

#### 3.5. Asignación Atómica de Nº ORDEN y Algoritmo de Búsqueda de Huecos
La arquitectura objetivo aprobada (Fase 6 / Rev. 5) sustituye definitivamente el cálculo por `MAX+1` por el algoritmo `BuscarSiguienteOrdenCandidataLibreBD`, bajo las siguientes reglas consensuadas:
1. **Permisión de Celdas Vacías Intermedias:**  
   El algoritmo reconoce que pueden existir celdas vacías intermedias en la columna A por cancelaciones o limpiezas visuales; la presencia de celdas vacías intercaladas no provoca la interrupción ni el aborto del proceso.
2. **Validación Exhaustiva de Valores Ocupados:**  
   Examina todos los valores no vacíos de la columna A, verificando que correspondan a enteros válidos dentro del intervalo oficial. Cualquier contenido no numérico o dato corrupto es rechazado formalmente.
3. **Lectura Matricial Única en Memoria:**  
   Lee el rango `A2:AúltimaFila` en una matriz `Variant` en RAM en una única llamada COM.
4. **Construcción de Conjuntos y Límites:**  
   Construye `dictPresentes` y calcula los valores extremos `minEnHoja` y `maxEnHoja`.
5. **Consulta Batch Única a Access:**  
   Ejecuta una única consulta parametrizada a la base de datos:
   ```sql
   SELECT NumOrden FROM Ordenes 
   WHERE Residencia = '...' AND NumOrden BETWEEN minEnHoja AND maxEnHoja
   ```
   cargando todos los números ocupados en el diccionario `dictAccess`.
6. **Búsqueda Ascendente de Huecos:**  
   Examina secuencialmente los enteros entre `minEnHoja` y `maxEnHoja`. El primer número que no exista ni en `dictPresentes` ni en `dictAccess` es seleccionado como candidato.
7. **Continuación Superior Secuencial:**  
   Si no existen huecos libres en el intervalo ocupado, el candidato seleccionado es el valor de la última fila ocupada más uno (`valorUltimaFila + 1`).
8. **Reserva Atómica y Confirmación Explícita del Operador (Sin Reasignación Automática):**  
   La asignación de números de orden se rige estrictamente por el principio de confirmación expresa del operador, prohibiéndose de forma absoluta cualquier sustitución o reintento automático silencioso:
   1. El algoritmo calcula un número candidato libre evaluando de forma combinada los números presentes en la cuadrícula de la hoja y los números ocupados en la tabla `Ordenes` de Access (`dictAccess`).
   2. El número candidato calculado se **presenta formalmente al operador** en la interfaz para su aceptación.
   3. Si el operador acepta el candidato propuesto, el sistema intenta reservar atómicamente en Access **exactamente ese número aceptado**.
   4. La operación de persistencia no puede cambiar por iniciativa propia ni alterar el número confirmado.
   5. Si otro operador ha reservado ese mismo número en el intervalo transcurrido entre la presentación y la confirmación, el intento de inserción colisiona, se captura el error y la transacción se revierte limpiamente (`RollbackTrans`).
   6. Se informa de inmediato al operador mediante un aviso explícito indicando que el número ha dejado de estar disponible por colisión concurrente.
   7. El sistema recalcula un nuevo candidato libre y lo presenta nuevamente al operador.
   8. El operador debe **confirmar expresamente el nuevo candidato** antes de que el sistema intente reservarlo.
   *(Queda terminantemente prohibido reintentar automáticamente con el siguiente número libre o asignar de forma desatendida un identificador distinto al confirmado).*

#### 3.6. Segregación Canónica de Rangos Numéricos (1-899999 Oficial, 900001-999999 Provisional)
Queda formalmente establecida la segregación estricta de intervalos de numeración:
- **Rango Oficial:** Números enteros en el intervalo **`[1 .. 899.999]`**. Son emitidos y consolidados exclusivamente por la base de datos central Access en modo conectado.
- **Rango Provisional (SICO-N2):** Números enteros en el intervalo **`[900.001 .. 999.999]`**. Son generados localmente por la función `ObtenerSiguienteNumOrdenProvisional` durante contingencias offline.
- **Exclusión del Valor `900000`:**  
  El valor **`900.000`** está estrictamente excluido tanto del rango oficial como del rango provisional. No puede ser emitido como número oficial ni como número provisional.
- **Rechazo de Valores Superiores (`>= 1.000.000`):**  
  Cualquier valor numérico mayor o igual a `1.000.000` debe ser formalmente rechazado, impidiendo riesgos de desbordamiento en conversiones a tipo `Long` y garantizando la coherencia de formato.
- **Función `EsNumOrdenProvisional`:**  
  La validación de provisionalidad se define con comprobación acotada de ambos extremos:
  ```vb
  Public Function EsNumOrdenProvisional(ByVal numOrden As Long) As Boolean
      EsNumOrdenProvisional = (numOrden >= 900001 And numOrden <= 999999)
  End Function
  ```
- **Bloqueo en Entrada Online:**  
  Los procedimientos de captura online rechazan de plano la inserción de números provisionales en modo conectado.

#### 3.7. Reconciliación FIFO Segura y Trazabilidad de Escenarios
Para transferir de forma fiable los expedientes registrados localmente en contingencia hacia la base de datos central una vez restablecida la red, el subsistema FIFO (`ModuloSeguridadLocal.bas:906-1116` y `modDatabase.bas:1737-1993`) procesa cronológicamente cada operación garantizando la integridad ante los siguientes tres escenarios de ejecución:

- **Escenario A (Registro Existe con Lease Lock Activo de Otro Puesto):**  
  Al ejecutar el `UPDATE` de adquisición del bloqueo sobre `Ordenes`, si existe una fila cuyo `UsuarioModificacion` es no nulo, no vacío y distinto de `tokenMod`, `UltimaModificacion` no es nula y `DateDiff('s', UltimaModificacion, Now())` está entre 0 y 119 segundos, la sentencia afecta a 0 registros (`recsLock = 0`). La rutina consulta mediante `SELECT` el titular y la marca temporal, calcula los segundos restantes, ejecuta `cn.RollbackTrans`, muestra el mensaje modal genérico de conflicto y llama a `SincronizarOrdenHaciaExcel(ws, numOrden, fila, resTarget)`. La fila local se refresca desde Access, **descartando el cambio local no consolidado**. Retorna `False`. Esta misma rama también se ejecuta si la clave no existe; en ese caso el `SELECT` devuelve `EOF`, pero el mensaje sigue atribuyendo el conflicto a "otro usuario" porque no existe una rama diferenciada para ese resultado.
- **Escenario B (Registro Existe pero su Lease Lock ha Caducado — 120 segundos o más):**  
  Si la fila existe y `DateDiff('s', [UltimaModificacion], Now()) >= 120`, el predicado de caducidad habilita la adquisición. La sentencia puede afectar a 1 fila (`recsLock = 1`); a continuación, la función abre el Recordset de la clave y actualiza los campos de datos, manteniendo `tokenMod` como titular, renueva `UltimaModificacion` y confirma la transacción con `cn.CommitTrans`. Retorna `True`. La marca temporal futura no es caducidad: se trata separadamente en el escenario 3 de la tabla del apartado 3.4.
- **Escenario C (Clave inexistente en Access al invocar `SincronizarFilaHaciaAccess`):**  
  Si no existe una fila para `(NumOrden, Residencia)`, el `UPDATE` de adquisición no puede afectar a una fila (`recsLock = 0`). La función entra en la misma rama de rechazo descrita en el escenario A: el `SELECT` posterior devuelve `EOF`, no se distingue esa situación, se conserva el titular predeterminado `"otro usuario"`, se muestra el mensaje genérico de conflicto, se ejecuta `RollbackTrans`, se llama a `SincronizarOrdenHaciaExcel(ws, numOrden, fila, resTarget)` y la función termina con `False`. La firma real de la función solo recibe `ws`, `fila` y `resTarget`; no existe un argumento `permitirAlta`.

  El código contiene además el bloque `If rs.EOF Then rs.AddNew` después de que la adquisición haya devuelto `recsLock = 1` y se haya abierto el Recordset de la clave. Ese bloque no es la ruta de recuperación del caso anterior (`recsLock = 0`).

  **Helper independiente `AsegurarOrdenExiste`:** existe una rutina separada (`modDatabase.bas:1698-1732`) que consulta `COUNT(*)` y, si no encuentra fila, realiza un `INSERT` de metadatos iniciales mediante una llamada independiente. No se atribuye aquí una invocación automática desde `SincronizarFilaHaciaAccess`: para afirmar que un flujo de negocio ejecuta primero ese helper y después sincroniza la fila debe verificarse el llamador concreto. La existencia del helper no modifica el comportamiento del caso de clave inexistente cuando se invoca directamente `SincronizarFilaHaciaAccess`.

**Ruta alternativa de importación masiva (no protegida por Lease Lock):** `SincronizarLoteImportacionBD` no debe confundirse con la ruta individual anterior. Aunque abre una transacción explícita y la revierte si se produce un error, carga los registros de la residencia en un Recordset y actualiza los que encuentra por `NumOrden`, sin evaluar previamente `UsuarioModificacion` ni `UltimaModificacion`. Al escribir sobre un registro existente sustituye el titular por `tokenActual` y renueva la marca temporal. Por ello, una importación masiva concurrente puede sobrescribir una edición que otra instancia mantiene bajo Lease Lock. Esta limitación es un defecto As-Built independiente de la atomicidad interna de la transacción de lote (`DEF-FIFO-07`).

#### 3.8. Persistencia Granular de Tarifas (TAR-10) vs Persistencia Base de Tarifas
- **Persistencia Base de Tarifas (Existente en Producción y Versión Referencia):**  
  La base de datos Access cuenta con la tabla `ConfigTarifasCeldas` (10 campos, 291 registros activos), donde se almacenan celda a celda los valores de `TARIFAS UNIDAD` y `Turnos-Precios` (`Hoja28`). El procedimiento `CargarTarifasDesdeBD` carga estas tarifas al abrir el libro e incluye la optimización de rendimiento (`ScreenUpdating = False` / `Calculation = xlCalculationManual`) y el blindaje contra el rango combinado `B25:D25`. Esto se encuentra plenamente operativo en producción y replicado en la Versión Referencia.
- **Mejora de Concurrencia Granular TAR-10 (Diseño Aprobado No Implementado):**  
  La propuesta TAR-10 contempla el guardado celda a celda con detección optimista de colisiones concurrentes (*Lost Updates*) y recarga diferencial sin refresco total. Dicha mejora **no está implementada en la Versión Referencia ni en el binario de laboratorio Fase 4**: constituye un diseño documental aprobado pendiente de desarrollo técnico e integración formal.

#### 3.9. Estado de Aprobación vs. Propuestas Pendientes
| Componente / Solución | Nivel de Certeza | Estado Formal | Entorno Actual | Pendiente de Inyección |
| :--- | :---: | :---: | :---: | :---: |
| **Persistencia Base Tarifas (291 celdas)** | Certeza 1 | **Operativo As-Built** | Producción (`Residencia_BE.accdb`) | No (Ya operativo) |
| **Persistencia Granular TAR-10** | Certeza 3 | **Diseño Aprobado No Implementado** | Diseño Documental (no en Ref ni Lab) | Sí (desarrollo e inyección) |
| **Blindaje Celdas Combinadas B25:D25**| Certeza 1 | **Operativo As-Built** | Producción (`modDatabase.bas`) | No (Ya operativo) |
| **Clase `clsDatabaseEngine` V5.2** | Certeza 4 | **Código Integrado en Lab (Homologación Pendiente de Evidencia)** | Binario Lab Fase 4 (`.xlsm`) | Sí (validación e inyección) |
| **CapKey de 256 bits en Fachada** | Certeza 4 | **Código Integrado en Lab (Homologación Pendiente de Evidencia)**| Binario Lab Fase 4 (`.xlsm`) | Sí (validación e inyección) |
| **Token de Sesión (`usuario@PC#HexGUID`)**| Certeza 1 | **Operativo As-Built (con limitaciones)**| Producción (`modDatabase.bas`) | No (Ya operativo) |
| **Búsqueda de Huecos ORD-10 (Rev. 5)** | Certeza 3 | **Diseño Aprobado No Integrado** | Diseño Documental Fase 6 | Sí (hacia Producción) |
| **Segregación Canónica de Rangos** | Certeza 3 | **Diseño Aprobado Pendiente** | Normativa Proyecto | Sí (homogeneización en código) |
| **Reconciliación FIFO Atómica (Rev. 5)**| Certeza 3 | **Diseño Aprobado No Integrado** | Propuesta Rev. 5 | Sí (implementación completa) |

---

### 4. BLOQUE III — CATÁLOGO DE DEFECTOS, RIESGOS Y EVOLUCIÓN PENDIENTE

#### 4.1. Fichas Técnicas de los 18 Defectos Estables

##### `DEF-FIFO-01` — Asignación de Número Oficial en Reconciliación mediante `MAX(NumOrden)+1`
- **Severidad:** Crítica.
- **Componentes y Líneas Implicadas:** `ModuloSeguridadLocal.bas:1364-1402`, procedimiento `ResolverNumOrdenProvisionalEnBD`.
- **Descripción Técnica del Fallo:** Durante la reconciliación de la cola de contingencia desconectada, el procedimiento calcula el nuevo número oficial ejecutando `SELECT MAX(NumOrden) FROM Ordenes WHERE Residencia = '...' AND NumOrden < 900001` (línea 1373) y le suma 1 en memoria sin reservar el número en Access.
- **Escenario de Reproducción:** Si dos terminales salen de contingencia simultáneamente o un operador online registra una reserva entre la ejecución del `SELECT MAX` y la posterior inserción, ambos reciben el mismo número oficial.
- **Impacto Operativo:** Error de colisión de clave única (`idx_Residencia_NumOrden`) o denegación de reconciliación de la reserva.
- **Decisión Arquitectónica Aprobada (ORD-10 Rev. 5):** Sustituir la lógica por invocación a `BuscarSiguienteOrdenCandidataLibreBD`, realizando búsqueda de huecos y reserva atómica en la misma transacción.

##### `DEF-FIFO-02` — Ausencia de Validación de Identidad de Expediente en `AsegurarOrdenExiste`
- **Severidad:** Crítica.
- **Componentes y Líneas Implicadas:** `ModuloSeguridadLocal.bas:1381`, llamada a `modDatabase.AsegurarOrdenExiste` (`modDatabase.bas:1698-1732`).
- **Descripción Técnica del Fallo:** El fallo no radica únicamente en una condición de carrera concurrente, sino en un **defecto de validación de identidad de expediente**. Antes de consolidar el número oficial, el código invoca `AsegurarOrdenExiste(nuevoNum, resCode)`. En `modDatabase.bas:1713-1716`, la función ejecuta `SELECT COUNT(*) FROM Ordenes WHERE NumOrden = ...`. Si el número ya existe en Access (`CLng(varCount) > 0`), la función retorna inmediatamente `True` asumiendo que el registro pertenece a la reserva que se está procesando, **sin realizar ninguna verificación de identidad material del expediente (no coteja DNI, Nombre de huésped ni Id autonumérico inmutable)**.
- **Escenario de Reproducción:** Una reserva en contingencia recibe durante la reconciliación un número oficial que ya había sido tomado y rellenado con datos de otro huésped por un puesto online. `AsegurarOrdenExiste` evalúa que el número existe, retorna `True` y valida erróneamente la sustitución.
- **Impacto Operativo:** Validación errónea de ocupación ajena como propia, habilitando la invasión destructiva de expedientes preexistentes.
- **Alternativa Propuesta Pendiente de Aprobación:** Validar la correspondencia material de negocio del expediente mediante clave subrogada inmutable (`Id` autonumérico o coincidencia de datos de reserva) y exigir reserva atómica en estado libre antes de la asignación.

##### `DEF-FIFO-03` — Sobreescritura Destructiva de Expedientes por Ausencia de Verificación de Identidad
- **Severidad:** Catastrófica.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:1737-1993` (`SincronizarFilaHaciaAccess`), invocado desde `ModuloSeguridadLocal.bas:1042-1048`.
- **Descripción Técnica del Fallo:** Si bien el algoritmo Lease Lock protege la edición temporal frente a modificaciones simultáneas durante 120 segundos, **no valida la identidad material del expediente**. Cuando `AsegurarOrdenExiste` valida un número ya existente en Access, `SincronizarFilaHaciaAccess` comprueba si el registro está bloqueado; si el bloqueo expiró o está libre, adquiere el Lease Lock y ejecuta un `UPDATE` masivo de campos sin comprobar si el DNI o el huésped existente coinciden con los que se intentan volcar.
- **Escenario de Reproducción:** Reconciliación de una reserva offline cuyo número calculado colisiona con un expediente preexistente en Access que no tiene un Lease Lock activo en ese instante.
- **Impacto Operativo:** Sobreescritura destructiva irreversible de la reserva original en Access, destruyendo los datos del huésped previo y su facturación.
- **Solución Arquitectónica Objetivo:** Validación cruzada de identidad de expediente (`Id` autonumérico inmutable o coincidencia estricta de DNI) previo a la actualización, impidiendo mutar registros cuyo origen no sea la reserva local que se reconcilia.

##### `DEF-FIFO-04` — Pérdida de Cambios Locales en Excel por Rechazo de Lease Lock y Refresco desde Servidor
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:1737-1993` (`SincronizarFilaHaciaAccess`, líneas 1834-1839), `ModuloSeguridadLocal.bas:906-1116` y `modDatabase.bas:3213-3500` (`SincronizarHojaDesdeAccess`).
- **Descripción Técnica del Fallo:** Comportamiento fáctico demostrado en código: cuando un operador intenta sincronizar una fila y el Lease Lock es rechazado porque otro puesto mantiene el expediente bloqueado, el sistema **no actúa de manera silenciosa**: muestra una **advertencia modal explícita** (`MsgBox` con icono `vbExclamation`, líneas 1834-1839) informando de qué usuario bloquea el expediente, cuántos segundos restan y advirtiendo de que la hoja se actualizará con los datos del servidor.  
  Sin embargo, a continuación invoca inmediatamente `SincronizarHojaDesdeAccess`, la cual sobrescribe las celdas de la cuadrícula con el contenido de Access. **Las modificaciones locales no consolidadas se pierden de forma irreversible porque ResidenciaApp no implementa ningún mecanismo de aislamiento, cola de retención, stash local ni resolución de conflictos**.
- **Escenario de Reproducción:** Un operador local edita una orden en Excel mientras otro terminal mantiene un bloqueo activo en Access; al guardar o perder el foco, el sistema muestra el mensaje de bloqueo y, al cerrar el diálogo, el refresco desde el servidor borra los cambios recién tecleados.
- **Impacto Operativo:** Pérdida definitiva del trabajo introducido por el usuario local, forzándolo a reintroducir todos los datos una vez expire el bloqueo del otro puesto.
- **Alternativa Propuesta Pendiente de Aprobación:** Implementar una capa de almacenamiento temporal o bifurcación local (stash de conflicto) que conserve los valores editados localmente y permita reintentar la consolidación o comparar diferencias tras liberarse el Lease Lock central.

##### `DEF-FIFO-05` — Creación de Registro Esqueleto en Access por Falta de Transacción Compartida
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:1698-1732` (`AsegurarOrdenExiste`) y `modDatabase.bas:1737-1993` (`SincronizarFilaHaciaAccess`).
- **Descripción Técnica del Fallo:** El procedimiento `AsegurarOrdenExiste` ejecuta la sentencia:
  ```sql
  INSERT INTO Ordenes (NumOrden, Residencia, FechaCreacion, UsuarioCreacion, 
                       UsuarioModificacion, UltimaModificacion, Solapamiento, ConsentimientoRGPD) 
  VALUES (<numOrden>, '<resTarget>', #Now#, '<userCreacion>', '<tokenActual>', #Now#, False, False)
  ```
  y la consolida en Access mediante `ExecuteNonQuery` de forma totalmente autónoma y desvinculada de la transacción posterior. Si la posterior llamada a `SincronizarFilaHaciaAccess` fracasa (por caída de red, colisión concurrente o error en validación de datos), **el registro inicial (con solo los 8 campos de cabecera y el resto nulos) queda consolidado permanentemente en Access sin compartir transacción con el llenado de datos**.
- **Escenario de Reproducción:** Creación de orden donde la inserción básica de 8 campos prospera en `AsegurarOrdenExiste`, pero el posterior volcado en `SincronizarFilaHaciaAccess` es interrumpido.
- **Impacto Operativo:** Existencia de órdenes esqueleto incompletas en Access que consumen números de orden e inducen a errores operativos.
- **Alternativa Propuesta Pendiente de Aprobación:** Integración de la comprobación, reserva y volcado completo de datos bajo un único bloque transaccional indivisible, o compensación atómica automática (eliminación del esqueleto si fracasa la sincronización posterior).

##### `DEF-FIFO-06` — Inconsistencia en Reintentos WAL tras Sustitución Prematura de Número
- **Severidad:** Media-Alta.
- **Componentes y Líneas Implicadas:** `ModuloSeguridadLocal.bas:1387-1395` (`ResolverNumOrdenProvisionalEnBD`).
- **Descripción Técnica del Fallo:** El sistema muta la columna de número de orden en la hoja oculta `SYS_COLA_SYNC` asignándole el nuevo número oficial antes de que la persistencia completa en Access haya sido confirmada definitivamente. Si el volcado posterior falla, los reintentos siguientes del WAL intentan reconciliar un número oficial como si fuera provisional.
- **Escenario de Reproducción:** Reintentos automáticos tras error intermitente de socket OLEDB.
- **Impacto Operativo:** Corrupción de la cola local de sincronización, requiriendo intervención técnica manual.
- **Solución Arquitectónica Objetivo:** Mantener inmutable el identificador provisional en el registro WAL hasta la confirmación definitiva (`Commit`) de la transacción central.

##### `DEF-FIFO-07` — Importación Masiva de Órdenes sin Adquisición del Lease Lock por Expediente
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:2021-2238` (`SincronizarLoteImportacionBD`). El inventario previo atribuye su invocación a `CopiarDatosHojasGijonSotoOviedo.bas`; esa relación de llamada no se revalidó en la Orden 1H.
- **Descripción Técnica del Fallo:** La rutina inicia una transacción (`cn.BeginTrans`) y procesa filas mediante un Recordset abierto sobre todas las órdenes de la residencia. Para cada número existente selecciona el registro por `NumOrden` y ejecuta `rs.Update`, pero no realiza un `UPDATE` condicional de adquisición del Lease Lock ni comprueba el titular/antigüedad de `UsuarioModificacion` y `UltimaModificacion`. En su lugar, escribe directamente el token de la instancia que ejecuta el lote y una marca temporal nueva. La transacción protege la consistencia interna del lote frente a errores de ejecución, pero no impide que la importación sobrescriba un expediente que otra instancia mantiene bloqueado mediante la ruta individual.
- **Escenario de Reproducción:** Un puesto mantiene una orden existente bajo un Lease Lock vigente. Otro puesto ejecuta una importación masiva para la misma residencia que incluye ese número de orden. La rutina por lotes carga y actualiza el registro sin comprobar el bloqueo, sustituyendo el token del titular y los datos del expediente.
- **Impacto Operativo:** Pérdida de cambios concurrentes y anulación práctica del aislamiento optimista proporcionado por `SincronizarFilaHaciaAccess` cuando se utiliza la ruta de importación masiva.
- **Corrección Pendiente de Decisión Arquitectónica:** La ruta por lotes deberá respetar el mismo contrato de concurrencia por expediente o rechazar de forma explícita las filas cuyo Lease Lock pertenezca a otra instancia. No se modifica el código en esta rectificación documental.

##### `DEF-NUM-01` — Ausencia de Reutilización de Huecos Interiores y Vulnerabilidad en Frontera 900000
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:1403-1472` (`ReservarSiguienteOrdenLibreBD`) y `modDatabase.bas:1527-1696` (`GestionarEntradaNumOrdenHoja`).
- **Descripción Técnica del Fallo:**  
  1. *Ausencia de Búsqueda de Huecos:* La función calcula sistemáticamente `MAX(NumOrden) + 1`. Si existen números libres intercalados debido a cancelaciones o borrados previos, jamás se reutilizan, provocando el agotamiento artificial de la secuencia hacia cotas elevadas.
  2. *Vulnerabilidad en el Límite 900000:* La consulta SQL `WHERE NumOrden < 900001` permite que, cuando el máximo existente es 899999, el candidato resultante sea `900000`. Al carecer de validación de techo en 899999, el sistema intenta insertar y consolidar 900000 como orden oficial. Asimismo, la entrada manual en celda de `900000` no es reconocida como provisional por `EsNumOrdenProvisional` (al evaluar `>= 900001`), permitiendo su persistencia directa como orden oficial en modo online.
- **Escenario de Reproducción:** Numeración de expedientes cuando la serie alcanza 899.999 o introducción manual de 900000.
- **Impacto Operativo:** Invasión del valor separador 900000 y falta de compactación de la serie de órdenes.
- **Solución Arquitectónica Objetivo:** Incorporar el algoritmo atómico de búsqueda de huecos interiores (propuesta ORD-10 / Rev. 5), homogeneizar la cota superior oficial en `899.999` y rechazar formalmente el valor `900.000` en todos los flujos de captura.

##### `DEF-NUM-02` — Ausencia de Cota Superior en Validación Provisional (`>= 1000000`)
- **Severidad:** Media.
- **Componentes y Líneas Implicadas:** `ModuloSeguridadLocal.bas:617-619`, función `EsNumOrdenProvisional`.
- **Descripción Técnica del Fallo:** La función `EsNumOrdenProvisional` evalúa exclusivamente `(numOrden >= MIN_ORDEN_PROVISIONAL)`, careciendo de comprobación de cota superior.
- **Escenario de Reproducción:** Entrada accidental o procesado de números de 7 o más dígitos (por ejemplo `1234567`).
- **Impacto Operativo:** Valores arbitrariamente grandes son clasificados erróneamente como provisionales de contingencia SICO-N2.
- **Solución Arquitectónica Objetivo:** Incorporar la condición acotada `(numOrden >= 900001 And numOrden <= 999999)` y rechazar formalmente cualquier entrada `>= 1000000`.

##### `DEF-NUM-03` — Desbordamiento en Conversión a Tipo `Long` (Error 6 Overflow)
- **Severidad:** Media.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:1548` (`GestionarEntradaNumOrdenHoja`), `DuplicadosOrdenyFactura.bas:124` (`VerificarSecuenciaOrden`) y eventos `Worksheet_Change` en `Hoja2.cls`, `Hoja20.cls`, `Hoja24.cls`.
- **Descripción Técnica del Fallo:** En `modDatabase.bas:1548`, el código ejecuta directamente `numSolicitado = CLng(valorCelda)` asumiendo que el contenido es convertible a entero largo de 32 bits. Si un usuario introduce o pega en la celda un número superior a $2.147.483.647$ (como un número de teléfono, IBAN o código de barras), VBA genera un **`Error 6: Desbordamiento (Overflow)`** no interceptado que aborta la macro y deja la cuadrícula en estado inconsistente. Idéntica vulnerabilidad existe en `DuplicadosOrdenyFactura.bas:124`.
- **Escenario de Reproducción:** Pegado de una cadena numérica extensa o número de teléfono en la celda de Nº ORDEN (Columna A).
- **Impacto Operativo:** Caída no controlada de la ejecución de macros de Excel con apertura de cuadros de diálogo del depurador ante el operador.
- **Solución Arquitectónica Objetivo:** Validación previa mediante comprobación de cadena (`If Len(valorStr) > 6 Or Not IsNumeric(valorStr) Then Exit Sub`) antes de invocar `CLng()`, limitando el rango aceptable a $[1 .. 999.999]$.

##### `DEF-NUM-04` — Falta de Segregación de Números Provisionales en Entrada Online
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** Controladores de eventos `Worksheet_Change` en `Hoja2.cls`, `Hoja20.cls`, `Hoja24.cls`.
- **Descripción Técnica del Fallo:** En modo conectado, la cuadrícula permite teclear manualmente un número en el rango provisional (`900001..999999`) sin verificar si la base de datos está disponible.
- **Escenario de Reproducción:** Entrada manual de un número en el rango de 900.000 por confusión del operador.
- **Impacto Operativo:** Inserción en la base de datos central de números provisionales fuera del flujo de contingencia.
- **Solución Arquitectónica Objetivo:** Rechazo absoluto de números provisionales en modo conectado; todo expediente debe recibir un número oficial `[1 .. 899.999]`.

##### `DEF-HIST-01` — Riesgo de Sustitución Destructiva de Históricos Previos al Reutilizar Números
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:4261-4420` (`EliminarOrdenBD`) y `modDatabase.bas:4424-4552` (`EliminarOrdenesLoteBD`).
- **Descripción Técnica del Fallo:** La tabla `Ordenes_Historico` carece de clave primaria y su índice `idx_Hist_Res_NumOrden` no es único (`Unique=False`). Sin embargo, los procedimientos de archivado aplican por lógica de aplicación una **política de depuración o sustitución previa**:
  ```sql
  DELETE FROM Ordenes_Historico WHERE NumOrden = N AND Residencia = '...'
  ```
  Si en el pasado se reasignó un número de orden o se reutilizó un identificador numérico antiguo en la cuadrícula activa, **este borrado previo destruye físicamente el registro histórico legítimo anterior que portaba ese mismo número**, sustituyéndolo de forma irreversible por la nueva orden que se archiva.
- **Escenario de Reproducción:** Archivado de una orden activa cuyo número coincide con un expediente histórico archivado en ejercicios o turnos pasados.
- **Impacto Operativo:** Pérdida silenciosa e irrecuperable de datos históricos y registros de facturación anteriores sin advertencia al usuario.
- **Alternativa Propuesta Pendiente de Aprobación:** Modificar la lógica de aplicación para que no elimine registros históricos previos, o dotar a `Ordenes_Historico` de un identificador autonumérico propio (`IdHistorico`) con marcas de versión, preservando todos los registros históricos generados.

##### `DEF-HIST-02` — Traspaso Masivo Involuntario en Vaciado Administrativo de Sede
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:4556-4659` (`EliminarTodasOrdenesSedeBD`).
- **Descripción Técnica del Fallo:** El procedimiento implementa estrictamente el Caso B (ejecuta sustitución previa en `Ordenes_Historico` solo de los números coincidentes con órdenes activas con datos reales, traspasa las activas al histórico y elimina de `Ordenes` las órdenes de esa sede). Sin embargo, carece de barreras reforzadas de confirmación en la capa de persistencia: aunque valida la ausencia de Lease Locks activos de otros puestos, una invocación indebida desencadena el traspaso masivo de todas las reservas de la sede al histórico.
- **Escenario de Reproducción:** Invocación inadvertida del procedimiento administrativo de vaciado de sede.
- **Impacto Operativo:** Traspaso no planificado de todas las reservas activas de la residencia al archivo histórico.
- **Alternativa Propuesta Pendiente de Aprobación:** Exigir credenciales de rol `ADMIN`, confirmación explícita de seguridad en la interfaz de usuario y copia snapshot previa antes de ejecutar el vaciado.

##### `DEF-ARCH-01` — Asimetría entre Persistencia Transaccional en Access y Limpieza Visual en Excel
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `modDatabase.bas:4784-4970` (`ArchivarOrdenesAntiguasBD`).
- **Descripción Técnica del Fallo:** El procedimiento opera en dos fases formalmente disjuntas y no atómicas entre sí: la Fase 1 ejecuta el traspaso y borrado en Access de forma atómica delegando en `EliminarOrdenesLoteBD`; posteriormente, la Fase 2 desprotege la hoja y ejecuta la compactación visual de filas en Excel (`Application.Union(...).Delete`). Si la Fase 2 se interrumpe por un error en Excel (como una celda bloqueada, fallo de memoria o desprotección fallida), la persistencia en Access ya fue confirmada irrevocablemente con `CommitTrans`, pero las filas archivadas permanecen visibles en la hoja de cálculo.
- **Escenario de Reproducción:** Interrupción por error de interfaz o fallo en la desprotección de la hoja durante la Fase 2 tras completarse con éxito la Fase 1 en Access.
- **Impacto Operativo:** El operador continúa viendo en pantalla órdenes que ya han sido eliminadas de la tabla activa `Ordenes`, generando intentos de edición sobre registros inexistentes.
- **Alternativa Propuesta Pendiente de Aprobación:** Orquestación determinista con verificación posterior de coherencia visual frente al estado del backend y rutina de sincronización visual forzada en el manejador de excepciones.

##### `DEF-TAR-01` — Ausencia de Inyección de la Mejora TAR-10 en Código Ejecutable
- **Severidad:** Media-Alta.
- **Componentes y Líneas Implicadas:** `bdas-multiusuario\BDAS_Multiusuario.xlsm`.
- **Descripción Técnica del Fallo:** Aunque la base de datos Access dispone de la tabla `ConfigTarifasCeldas`, la optimización de persistencia granular TAR-10 para la gestión de colisiones concurrentes celda a celda continúa en estado de diseño y no ha sido inyectada en el binario activo de producción.
- **Escenario de Reproducción:** Edición concurrente de tarifas por dos administradores simultáneamente.
- **Impacto Operativo:** Posibilidad de sobreescritura de valores tarifarios si dos sesiones modifican la misma celda.
- **Solución Arquitectónica Objetivo:** Inyección formal de los controladores saneados de TAR-10 desde la Versión Referencia hacia el binario maestro de producción.

##### `DEF-ORD-01` — Falta de Reutilización de Huecos Interiores y Dependencia de `MAX+1`
- **Severidad:** Alta.
- **Componentes y Líneas Implicadas:** `bdas-multiusuario\BDAS_Multiusuario.xlsm`, `modDatabase.bas:1403-1472` (`ReservarSiguienteOrdenLibreBD`).
- **Descripción Técnica del Fallo:** En Access, el índice único compuesto `idx_Residencia_NumOrden` sobre `(Residencia, NumOrden)` **previene físicamente la existencia de órdenes duplicadas en la base de datos**, provocando un Error 3022 si se intenta insertar una clave ya existente. Sin embargo, el defecto arquitectónico radica en que `ReservarSiguienteOrdenLibreBD` calcula candidatos exclusivamente mediante `MAX(NumOrden) + 1`: 1) No reutiliza huecos interiores liberados por cancelaciones o archivados; 2) Ante reservas simultáneas, puede generar colisiones de candidatos y reintenta con un límite de cinco intentos (`maxIntentos = 5`, `modDatabase.bas:1425-1429`); 3) El algoritmo ORD-10 de búsqueda de huecos aprobado en la Fase 6 (Rev. 5) permanece sin inyectar en el código de producción.
- **Escenario de Reproducción:** Cancelación de expedientes que dejan huecos en la serie numérica o alta simultánea de reservas por dos terminales.
- **Impacto Operativo:** Crecimiento acelerado e irreversible de la numeración hacia la barrera de 900.000, contención en reintentos y agotamiento prematuro del rango oficial.
- **Decisión Arquitectónica Aprobada (ORD-10 Rev. 5):** Inyección formal en producción del algoritmo ORD-10 (`BuscarSiguienteOrdenCandidataLibreBD`), que reutiliza huecos interiores y presenta al operador el candidato para confirmación expresa antes de la reserva atómica, sin reasignación automática ante colisiones.

##### `DEF-REPO-01` — Discrepancia entre Componentes del Binario y Archivos Exportados en Repositorio
- **Severidad:** Media.
- **Componentes y Líneas Implicadas:** Directorio `H:\ResidenciaApp\bdas-multiusuario\vba-modules\`.
- **Descripción Técnica del Fallo:** Coexistencia en la carpeta de exportación de 198 archivos físicos frente a los 173 componentes activos del proyecto binario compilado (incluyendo los 8 archivos huérfanos acreditados).
- **Escenario de Reproducción:** Auditoría de coherencia entre el repositorio Git y el binario Excel.
- **Impacto Operativo:** Confusión en tareas de mantenimiento técnico sobre qué archivos reflejan la versión activa en producción.
- **Solución Arquitectónica Objetivo:** Aislamiento y catalogación formal de los archivos huérfanos en el Documento Maestro, preservando su inmutabilidad según las directrices de auditoría.

##### `DEF-REST-01` — Scripts de Restauración sin Verificación Criptográfica de Hashes
- **Severidad:** Media.
- **Componentes y Líneas Implicadas:** Scripts PowerShell en `scripts\Restaurar-Version-Referencia.ps1` y `scripts\Restaurar-Golden-Master.ps1`.
- **Descripción Técnica del Fallo:** Los scripts de restauración efectúan copias de archivos basándose exclusivamente en rutas y nombres, sin verificar el hash SHA-256 de origen ni el de destino tras la operación.
- **Escenario de Reproducción:** Restauración tras incidente en caso de corrupción silenciosa en las plantillas.
- **Impacto Operativo:** Riesgo de propagar réplicas corruptas o desalineadas respecto al hito histórico homologado.
- **Alternativa Propuesta Pendiente de Aprobación:** Incorporar validación obligatoria de hash SHA-256 en los scripts de restauración antes y después de sobreescribir los archivos de trabajo.

---

### 5. MATRIZ MAESTRA DE REQUISITOS DEL SISTEMA

Se consolidan los requisitos del sistema categorizados por su origen y estado formal de vigencia.

> [!NOTE] Criterio de Certeza Documental en la Matriz
> En cumplimiento estricto de las directrices de auditoría, aquellos requisitos de negocio cuya formulación inicial no conste explícitamente en actas o especificaciones fundacionales históricas del proyecto se clasifican con el origen documental: **"Requisito funcional identificado; origen histórico pendiente de confirmar"**, evitando conjeturas retrospectivas.

#### 5.1. Categoría A: Requisitos del Negocio
| ID Requisito | Descripción Funcional | Ámbito / Sede | Regla de Negocio | Implementación As-Built | Origen Documental | Estado | Incidencias |
| :--- | :--- | :---: | :--- | :--- | :--- | :---: | :---: |
| **`REQ-ORIG-01`** | Registro de órdenes de alojamiento | GIJÓN, SOTO, OVIEDO | Cada reserva ocupa una fila estructurada en Excel (28 col en GIJÓN/SOTO, 33 en OVIEDO) y 43 campos en base de datos. | Hoja Excel + tabla `Ordenes`. | Requisito funcional identificado; origen histórico pendiente de confirmar | **Vigente As-Built** | Ninguna |
| **`REQ-ORIG-02`** | Generación de facturas correlativas | Todas las sedes | Numeración secuencial anual correlativa por residencia (en producción As-Built susceptible de huecos si se anula una factura o aborta la transacción). | `ContadorFacturas` en Access. | Requisito funcional identificado; origen histórico pendiente de confirmar | **Vigente As-Built** | Ninguna |
| **`REQ-ORIG-03`** | Visualización en calendario | Todas las sedes | Matriz de habitaciones vs días con código de colores. | Hojas `Calendario GIJÓN/SOTO/OVIEDO`. | Requisito funcional identificado; origen histórico pendiente de confirmar | **Vigente As-Built** | Ninguna |
| **`REQ-ORIG-04`** | Control de lista negra | Todas las sedes | Denegación de reserva ante coincidencias de DNI o Nombre. | Tabla `ListaNegra` + formulario `frmGestionListaNegra`. | Requisito funcional identificado; origen histórico pendiente de confirmar | **Vigente As-Built** | Ninguna |
| **`REQ-ORIG-05`** | Bloqueo temporal de habitaciones | Todas las sedes | Inhabilitar celdas por avería, mantenimiento u obras. | Tabla `BloqueosHabitaciones` + formulario. | Requisito funcional identificado; origen histórico pendiente de confirmar | **Vigente As-Built** | Ninguna |

#### 5.2. Categoría B: Requisitos Incorporados en Fases de Desarrollo
| ID Requisito | Descripción Funcional | Ámbito / Sede | Regla de Negocio | Implementación As-Built | Origen Documental | Estado | Incidencias |
| :--- | :--- | :---: | :--- | :--- | :--- | :---: | :---: |
| **`REQ-DEV-01`** | Contingencia desconectada (SICO-N2) | Todas las sedes | Permitir registrar reservas si la red cae temporalmente. | `SYS_COLA_SYNC` + archivo local `.wal`. | Fase 2 (Especificación SICO-N2) | **Vigente As-Built** | `DEF-FIFO-01..06`, `DEF-NUM-02` |
| **`REQ-DEV-02`** | Numeración provisional offline | Todas las sedes | Asignar números provisionales `>= 900001` sin red examinando la hoja activa. | `ObtenerSiguienteNumOrdenProvisional`. | Fase 2 (Especificación SICO-N2) | **Vigente As-Built** | `DEF-NUM-01`, `DEF-NUM-02` |
| **`REQ-DEV-03`** | Control optimista Lease Lock | Maestro Órdenes | Arrendamiento de edición de 120 s mediante token `usuario@PC#HexGUID`. | `UPDATE` condicional en `modDatabase.bas`. | Fase 3 (Control Concurrencia) | **Vigente As-Built** | `DEF-FIFO-03`, `DEF-FIFO-04` |
| **`REQ-DEV-04`** | Persistencia base de tarifas | Todas las sedes | Sincronización celda a celda de tarifas en Access. | Tabla `ConfigTarifasCeldas` (291 celdas). | Hito Noviembre 2026 | **Vigente As-Built** | Ninguna |
| **`REQ-DEV-05`** | Purga de histórico invocable | Histórico | Purga física a demanda de registros con antigüedad superior al parámetro numérico `anosRetencion` (default en código: 3). | Procedimiento `PurgarHistoricoRGPD` invocado desde `frmPanelRGPD`. | Código As-Built de Producción | **Vigente As-Built** | Ninguna |

#### 5.3. Categoría C: Requisitos Aprobados para Arquitectura Objetivo
| ID Requisito | Descripción Funcional | Ámbito / Sede | Regla de Negocio | Especificación Técnica | Origen Documental | Estado | Incidencias |
| :--- | :--- | :---: | :--- | :--- | :--- | :---: | :---: |
| **`REQ-OBJ-01`** | Búsqueda y reserva de huecos ORD-10 | Todas las sedes | Búsqueda del menor entero libre entre `minEnHoja` y `maxEnHoja` con reserva atómica. | `BuscarSiguienteOrdenCandidataLibreBD`. | Fase 6 (Rev. 5 Aprobada) | **Aprobado en Diseño** | `DEF-ORD-01`, `DEF-FIFO-01` |
| **`REQ-OBJ-02`** | Segregación canónica de rangos | Todas las sedes | Oficial `[1..899999]`, Excluido `900000`, Provisional `[900001..999999]`, Inválido `>=1000000`. | Modificación de validadores y conversiones. | Fase 6 (Auditoría 0B/0C) | **Aprobado en Diseño** | `DEF-NUM-01..04` |
| **`REQ-OBJ-03`** | Mediación transaccional con CapKey | Núcleo de Persistencia | Ninguna escritura a BD sin pasar por `clsDatabaseEngine` con token de 256 bits. | `ExecuteNonQuery` y métodos con CapKey. | Fase 4 (Arquitectura V5.2) | **Aprobado en Diseño; código presente en laboratorio; homologación funcional NO ACREDITADA (Certeza 4)** | Todas las de persistencia |
| **`REQ-OBJ-04`** | Reconciliación FIFO atómica | Sincronización SICO-N2 | Conciliación de expedientes locales con reserva atómica previa y Lease Lock. | Rediseño de `SincronizarColaOfflineHaciaAccess`.| Fase 6 (Propuesta Rev. 5) | **Aprobado en Diseño** | `DEF-FIFO-01..06` |
| **`REQ-OBJ-05`** | Persistencia granular TAR-10 | Configuración Tarifas | Control de colisiones optimista celda a celda en `ConfigTarifasCeldas`. | Módulos saneados TAR-10. | Diseño Documental Aprobado (no en Ref ni Lab) | **Aprobado en Diseño** | `DEF-TAR-01` |

#### 5.4. Categoría D: Mejoras Propuestas o Aplazadas
En cumplimiento de las instrucciones de la Orden 1E, se clasifican estas iniciativas expresamente como propuestas técnicas no aprobadas, sin atribuirles compromiso formal ni considerarlas requisitos aprobados del proyecto.

| ID Propuesta | Descripción Funcional | Justificación Técnica | Estado Formal |
| :--- | :--- | :--- | :---: |
| **`PROP-ARQ-01`** | Desacoplamiento del repositorio de archivos compartidos | Mantener la arquitectura cliente-servidor basada en archivos Access sobre SMB optimizando el control de bloqueos y tiempos de respuesta. | **Propuesta No Aprobada (Sin compromiso de proyecto)** |
| **`PROP-ARQ-02`** | Homogeneización formal de interfaces entre sedes | Evaluar la unificación de columnas secundarias entre GIJÓN, SOTO y OVIEDO sin alterar los campos específicos de negocio (`Turno` vs `Comision`, `CamaSuple`). | **Propuesta No Aprobada (Sin compromiso de proyecto)** |

---

### 6. MODELO FÍSICO DE DATOS DETALLADO (11 TABLAS DE USUARIO)

Se documenta la estructura física exacta de las **11 tablas de usuario reales** presentes en la base de datos de producción `H:\ResidenciaBD\Residencia_BE.accdb` (SHA-256: `ac2fb724...`), verificada directamente mediante inspección COM ADOX y OLEDB en la Orden 1C.

> [!IMPORTANT]
> **Certificación Negativa de Tablas Inexistentes:**
> Se refuta formalmente la existencia física de las tablas `HistorialAcceso`, `PreciosUnidad`, `RegistroLeaseLocks`, `SesionesActivas`, `TarifasUnidad` y `TurnosPrecios`. Ninguna de ellas existe en la base de datos de producción.
> La persistencia de tarifas se realiza íntegramente en `ConfigTarifasCeldas` (291 registros relacionales); los bloqueos de concurrencia se gestionan en los campos de `Ordenes`; y los eventos de auditoría residen en `LogActividad`.

#### 1. `Ordenes` (Tabla Maestra Transaccional)
- **Propósito:** Registro operativo principal de reservas y estancias en las tres residencias.
- **Campos (43):**
  - `Id`: Long (AutoIncrement). **Clave Primaria**.
  - `Residencia`: Text(20) [NOT NULL].
  - `NumOrden`: Long [NOT NULL].
  - `FechaCreacion`: DateTime.
  - `UsuarioCreacion`: Text(50).
  - `UsuarioModificacion`: Text(50) (Almacena el token de instancia `usuario@PC#HexGUID` para Lease Lock).
  - `UltimaModificacion`: DateTime (Marca temporal de adquisición/renovación de Lease Lock).
  - `FechaPeticion`: DateTime.
  - `NumFactura`: Long.
  - `Finalidad`: Text(50).
  - `Empleo`: Text(50).
  - `Situacion`: Text(50).
  - `Evaluacion`: Text(50).
  - `Turno`: Text(50) (Específico para Residencia Soto del Barco).
  - `Comision`: Text(50) (Específico para Residencias GIJÓN y OVIEDO).
  - `DNI`: Text(20).
  - `Rango`: Text(50).
  - `Nombre`: Text(100).
  - `Apellidos`: Text(100).
  - `FechaEntrada`: DateTime.
  - `FechaSalida`: DateTime.
  - `DiasUso`: Long.
  - `PAX`: Long.
  - `NumHabIndividuales`: Long.
  - `NumHabDobles`: Long.
  - `CamaSuple`: Text(50).
  - `Importe`: Double.
  - `TipoHuesped`: Text(30).
  - `EstadoPago`: Text(30).
  - `PagadoResidencia`: Text(30).
  - `Telefono`: Text(20).
  - `Email`: Text(100).
  - `Direccion`: Text(200).
  - `Poblacion`: Text(100).
  - `Provincia`: Text(50).
  - `CodigoPostal`: Text(10).
  - `Observaciones`: LongText (Memo).
  - `Resolucion`: Text(30).
  - `FechaGrabacion`: Text(255).
  - `HabitacionesAsignadas`: Text(200).
  - `Solapamiento`: Boolean.
  - `ConsentimientoRGPD`: Boolean.
  - `DtoFamNum`: Text(50).
- **Índices y Claves:**
  - Clave Primaria: `Id` (Unique=True, PrimaryKey=True).
  - Clave Única: `idx_Residencia_NumOrden` sobre los campos `(Residencia, NumOrden)` (Unique=True).
  - Índices Secundarios (No únicos): `idx_Ordenes_DNI`, `idx_Ordenes_FechaEntrada`, `idx_Ordenes_FechaSalida`, `idx_Ordenes_Residencia`, `idx_Ordenes_Resolucion`.

#### 2. `Ordenes_Historico` (Archivo Histórico de Expedientes)
- **Propósito:** Almacén de expedientes cerrados, cancelados o traspasados tras finalizar la estancia o por depuración de activos.
- **Campos (43):** Estructura idéntica a `Ordenes` (mismos 43 campos y tipos).
- **Índices y Claves:**
  - Clave Primaria: Ninguna (0 claves primarias definidas).
  - Índices: `idx_Hist_Res_NumOrden` sobre `(Residencia, NumOrden)` (Unique=False, para agilizar búsquedas por número y residencia).

#### 3. `ConfigTarifasCeldas` (Persistencia Base de Tarifas)
- **Propósito:** Almacenamiento celda a celda de los parámetros tarifarios de `TARIFAS UNIDAD` y `Turnos-Precios` (`Hoja28`) con 291 registros activos.
- **Campos (10):**
  - `Id`: Long (AutoIncrement). **Clave Primaria**.
  - `NombreHoja`: Text(40) [NOT NULL].
  - `Fila`: Long [NOT NULL].
  - `Columna`: Long [NOT NULL].
  - `TipoDato`: Text(10) (`TEXTO`, `NUMERO`, `FECHA`).
  - `ValorTexto`: Text(255).
  - `ValorNumero`: Double.
  - `ValorFecha`: DateTime.
  - `UsuarioModificacion`: Text(50).
  - `UltimaModificacion`: DateTime.
- **Índices y Claves:**
  - Clave Primaria: `Id` (Unique=True, PrimaryKey=True).
  - Clave Única: `idx_Tarifas_HojaCelda` sobre `(NombreHoja, Fila, Columna)` (Unique=True).

#### 4. `BloqueosHabitaciones` (Gestión de Indisponibilidad)
- **Propósito:** Registro de periodos en los que una habitación queda fuera de servicio por avería, mantenimiento u obras.
- **Campos (6):** `Id` (AutoInc, PK), `FechaDesde` (DateTime), `FechaHasta` (DateTime), `Habitacion` (Text 50), `Motivo` (Text 255), `Residencia` (Text 20).
- **Índices y Claves:**
  - Clave Primaria: `Id` (Unique=True, PrimaryKey=True).
  - Índice Secundario: `idx_Bloq_Res` sobre el campo `Residencia` (Unique=False).  
    *Aclaración de auditoría:* No existe ningún índice compuesto en esta tabla en el esquema físico real.

#### 5. `CodigosPostales` (Normalización de Poblaciones)
- **Propósito:** Maestro auxiliar para autocompletado de poblaciones y provincias al introducir el código postal.
- **Campos (5):** `Id` (AutoInc, PK), `CodigoPostal` (Text 10, NOT NULL), `Municipio` (Text 100, NOT NULL), `MunicipioNormalizado` (Text 100), `Provincia` (Text 50, NOT NULL).
- **Índices y Claves:**
  - Clave Primaria: `Id` (Unique=True, PrimaryKey=True).
  - Índices Secundarios: `idx_CP` (`CodigoPostal`), `idx_Municipio` (`Municipio`), `idx_MuniNorm` (`MunicipioNormalizado`), `idx_Provincia` (`Provincia`).

#### 6. `ContadorFacturas` (Serie Correlativa de Facturación)
- **Propósito:** Control concurrente de los contadores secuenciales de facturación por residencia y ejercicio anual.
- **Campos (4):** `Id` (AutoInc, PK), `Residencia` (Text 20, NOT NULL), `Ejercicio` (Long, NOT NULL), `UltimoNumero` (Long, NOT NULL).
- **Índices y Claves:**
  - Clave Primaria: `Id` (Unique=True, PrimaryKey=True).
  - Clave Única Compuesta: `idx_CF_ResEjer` sobre los campos `(Residencia, Ejercicio)` (Unique=True).

#### 7. `EstadoHabitacionesCalendario` (Visibilidad de Habitaciones en Calendario)
- **Propósito:** Control de visibilidad (filas mostradas u ocultas) de las habitaciones en las hojas de calendario operativo de cada sede.
- **Campos Físicos Reales (Exactamente 4 Columnas):**
  1. `Id`: Long (AutoIncrement). **Clave Primaria**. (Tipo OLEDB 3 / adInteger).
  2. `Residencia`: Text(20) [NOT NULL]. (Tipo OLEDB 202 / adVarWChar, DefinedSize: 20).
  3. `FilaCalendario`: Long [NOT NULL]. (Tipo OLEDB 3 / adInteger).
  4. `Oculta`: Boolean [NOT NULL]. (Tipo OLEDB 11 / adBoolean, DefinedSize: 2).
  *(Certificación de auditoría Orden 1C: Se constata mediante ADOX que la tabla posee exactamente 4 columnas. NO existe ninguna columna denominada `Habitacion`, asociándose el estado de visibilidad directamente al número de fila física del calendario).*
- **Índices y Claves Reales:**
  - Clave Primaria: `Index_7A884EDF_7C1C_4FC1` sobre `Id` (Unique=True, PrimaryKey=True).
  - Clave Única Compuesta: `idx_EstHab_ResFila` sobre los campos `(Residencia, FilaCalendario)` (Unique=True). (No existe ningún índice `idx_EHC_ResHab`).

#### 8. `ListaNegra` (Control de Restricciones de Acceso)
- **Propósito:** Registro de personas físicas con prohibición o restricción administrativa de alojamiento.
- **Campos Físicos Reales (6 Columnas):**
  1. `Id`: Long (AutoIncrement). **Clave Primaria**. (Tipo OLEDB 3 / adInteger).
  2. `DNI`: Text(20) [NOT NULL]. (Tipo OLEDB 202 / adVarWChar, DefinedSize: 20).
  3. `Nombre`: Text(150). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 150).
  4. `Motivo`: LongText / Memo. (Tipo OLEDB 203 / adLongVarWChar).
  5. `FechaAlta`: DateTime. (Tipo OLEDB 7 / adDate).
  6. `Activo`: Boolean. (Tipo OLEDB 11 / adBoolean, DefinedSize: 2).
  *(Certificación de auditoría: El campo de estado es `Activo` (Boolean) y el campo `Nombre` posee un ancho de 150 caracteres. No existe columna `UsuarioAlta`).*
- **Índices y Claves Reales:**
  - Clave Primaria: `Index_438A60E9_426F_4186` sobre `Id` (Unique=True, PrimaryKey=True).
  - Clave Única: `idx_LN_DNI` sobre el campo `DNI` (Unique=True).
  *(Certificación de índices: El índice sobre `DNI` es ÚNICO. No existe ningún índice `idx_LN_Nombre` en la base de datos).*
- **Consumo Fáctico en Código VBA (`modDatabase.bas:3114-3166`):**
  - Inserción/Reactivación (`AnadirAListaNegra`, líneas 3114-3155): Comprueba existencia por DNI; si ya existe ejecuta `UPDATE ListaNegra SET Nombre=..., Motivo=..., FechaAlta=Now(), Activo=True WHERE DNI=...`; si no existe ejecuta `INSERT INTO ListaNegra (DNI, Nombre, Motivo, FechaAlta, Activo) VALUES (...)`.
  - Retirada/Desactivación Lógica (`RetirarDeListaNegra`, líneas 3158-3166): Ejecuta `UPDATE ListaNegra SET Activo = False WHERE DNI = '...'`. **No ejecuta un `DELETE FROM`**, preservando el historial del expediente con baja lógica.

#### 9. `LogActividad` (Auditoría del Sistema)
- **Propósito:** Trazabilidad de operaciones críticas (inicios de sesión, cambios de tarifas, archivados y borrados).
- **Campos Físicos Reales (5 Columnas):**
  1. `Id`: Long (AutoIncrement). **Clave Primaria**. (Tipo OLEDB 3 / adInteger).
  2. `FechaHora`: DateTime. (Tipo OLEDB 7 / adDate).
  3. `Usuario`: Text(50). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 50).
  4. `Accion`: Text(200). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 200).
  5. `Detalle`: LongText / Memo. (Tipo OLEDB 203 / adLongVarWChar).
  *(Certificación de auditoría: Las columnas de descripción son `Accion` y `Detalle`. No existen columnas llamadas `Operacion` ni `DireccionIP`).*
- **Índices y Claves Reales:**
  - Clave Primaria: `Index_18335A05_CEAB_4709` sobre `Id` (Unique=True, PrimaryKey=True).
  - Índices Secundarios (No únicos): `idx_Log_FechaHora` sobre `FechaHora` e `idx_Log_Usuario` sobre `Usuario`.
- **Consumo Fáctico en Código VBA (`modDatabase.bas:3176-3201`):**
  - Registro de Auditoría (`InsertarLog`): Ejecuta exactamente `INSERT INTO LogActividad (FechaHora, Usuario, Accion, Detalle) VALUES (...)`, manteniendo una correspondencia del 100% campo a campo con el catálogo físico.

#### 10. `Usuarios` (Control de Acceso y Roles)
- **Propósito:** Gestión de operadores autorizados, claves criptográficas, roles y sedes permitidas.
- **Campos Físicos Reales (8 Columnas):**
  1. `Id`: Long (AutoIncrement). **Clave Primaria**. (Tipo OLEDB 3 / adInteger).
  2. `NombreUsuario`: Text(50) [NOT NULL]. (Tipo OLEDB 202 / adVarWChar, DefinedSize: 50).
  3. `Clave`: Text(128) [NOT NULL]. (Tipo OLEDB 202 / adVarWChar, DefinedSize: 128; almacena el hash SHA-256 de 64 caracteres).
  4. `NombreCompleto`: Text(100). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 100).
  5. `Rol`: Text(20). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 20; valores `ADMIN` o `OPERADOR`).
  6. `Residencias`: Text(60). (Tipo OLEDB 202 / adVarWChar, DefinedSize: 60; lista de sedes asignadas).
  7. `Activo`: Boolean. (Tipo OLEDB 11 / adBoolean, DefinedSize: 2).
  8. `FechaAlta`: DateTime. (Tipo OLEDB 7 / adDate).
- **Índices y Claves Reales:**
  - Clave Primaria: `Index_5FB45ACD_A5CD_4D5B` sobre `Id` (Unique=True, PrimaryKey=True).
  - Clave Única: `idx_Usr_NombreUsuario` sobre `NombreUsuario` (Unique=True).

#### 11. `BDAS_Historico_Snapshot` (Réplica Snapshot de Soporte)
- **Propósito:** Tabla de instantánea técnica auxiliar utilizada en operaciones de archivado masivo en `ModuloArchivadoDatosBDAS.bas`.
- **Campos (41):** `Apellidos`, `CamaSuple`, `CodigoPostal`, `Comision`, `ConsentimientoRGPD`, `DiasUso`, `Direccion`, `DNI`, `DtoFamNum`, `Email`, `Empleo`, `EstadoPago`, `Evaluacion`, `FechaCreacion`, `FechaEntrada`, `FechaGrabacion`, `FechaPeticion`, `FechaSalida`, `Finalidad`, `HabitacionesAsignadas`, `Id` (AutoInc), `Importe`, `Nombre`, `NumFactura`, `NumHabDobles`, `NumHabIndividuales`, `NumOrden`, `Observaciones`, `PagadoResidencia`, `PAX`, `Poblacion`, `Provincia`, `Rango`, `Residencia`, `Resolucion`, `Situacion`, `Solapamiento`, `Telefono`, `TipoHuesped`, `Turno`, `UsuarioCreacion`.
- **Índices y Claves:**
  - Clave Primaria: Ninguna (0 claves primarias definidas).
  - Índices: Ninguno (0 índices definidos).

---

> [!IMPORTANT] Ausencia Factual de Relaciones Físicas y Claves Foráneas
> La inspección exhaustiva de la colección `Keys` en el catálogo ADOX de `Residencia_BE.accdb` confirma de forma inequívoca que **no existe ninguna restricción de clave foránea física (Foreign Key) definida en el motor Access**. La integridad referencial del sistema depende al 100% de la lógica procedural codificada en VBA.

---

### 7. INVENTARIO DE COMPONENTES Y TRAZABILIDAD DE FUENTES

#### 7.1. Cifras Verificadas y Desglose de los 173 Componentes del Binario
La extracción y análisis del proyecto VBA de `BDAS_Multiusuario.xlsm` (SHA-256: `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`) confirma la existencia de **exactamente 173 componentes VBA en el binario activo**:
- **106 Módulos Estándar (`.bas`):** Lógica procedimental, conexión a base de datos (`modDatabase`), seguridad local (`ModuloSeguridadLocal`), cálculo de turnos y facturación.
- **49 Módulos de Hoja (`.cls`):** Controladores de hojas operativas (`Hoja2` GIJÓN, `Hoja20` SOTO, `Hoja24` OVIEDO), hoja de tarifas `Turnos-Precios` (`Hoja28`), calendarios y hojas ocultas del sistema.
- **1 Módulo `ThisWorkbook` (`.cls`):** Eventos de ciclo de vida del libro, control de guardado y barras de herramientas.
- **17 Formularios de Usuario (`.frm`):** Interfaz gráfica de inicio de sesión (`frmLogin`), bloqueos, listas negras y panel RGPD.
- **Total en Binario Activo:** **173 componentes VBA**.

> [!NOTE] Criterio de Verificación de Compilación
> Se hace constar con rigor técnico que la afirmación de ausencia de errores de compilación debe respaldarse mediante pruebas independientes de compilación dinámica en el entorno de ejecución de Microsoft Excel.

#### 7.2. Catálogo Acreditado de los 8 Archivos Huérfanos / Redundantes en Repositorio
En la carpeta `vba-modules/` del repositorio existen **8 archivos** que no están integrados en el proyecto binario compilado de producción `BDAS_Multiusuario.xlsm` (el cual contiene 173 componentes VBA acreditados):
1. `CalcularFacturaEstudiante.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
2. `modMigracion.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
3. `ModuloAjusteZoomHojas.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
4. `ModuloCrearHojaConfig.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
5. `ModuloGestionCacheCalendarios.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
6. `SolapesAceptadosListado.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
7. `TemporalEliminarLuego.bas`: Archivo de módulo estándar en repositorio, ausente en el binario compilado.
8. `Hoja34_TarifasUnidad.cls`: Archivo de módulo de clase en repositorio, ausente en el binario compilado.

*Regla de Inmutabilidad:* En cumplimiento estricto de las directrices del proyecto, estos 8 archivos permanecen intactos en su ubicación actual sin ser eliminados ni renombrados.

#### 7.3. Recurso RibbonX (`customUI/customUI.xml` en Paquete XLSM)
- **Ubicación en el Paquete OpenXML del Libro:**  
  La interfaz Ribbon personalizada del libro se encuentra almacenada en la ruta física interna:
  ```
  customUI/customUI.xml
  ```
- **Distinción de Fuentes:**  
  Debe diferenciarse el recurso integrado dentro del paquete ZIP del archivo `BDAS_Multiusuario.xlsm` de su posible réplica o archivo de trabajo en el árbol del repositorio.  
  *Corrección técnica de nomenclatura:* Queda expresamente desestimada la denominación `customUI14.xml` por no corresponder a la ruta real implementada en el contenedor del libro de producción.

#### 7.4. Trazabilidad Línea a Línea y Hashes Canónicos
- **Libro de Producción Activa:**  
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`  
  `SHA-256: b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7`
- **Base de Datos Access de Producción Activa:**  
  `H:\ResidenciaBD\Residencia_BE.accdb`  
  `SHA-256: ac2fb7240360b3771eca60f61ca8cf32fffeaf78e88b09d783c169a228bfe5cf`
- **Réplica Golden Master (Inmutable, Hito 2026-10-24):**  
  `H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm`  
  `H:\ResidenciaBD\Plantilla\GoldenMaster\Residencia_BE_GOLDEN_MASTER.accdb`
- **Réplica Versión Referencia (Inmutable, Hito 2026-11-10):**  
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`  
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`

---

### 8. REGLAS DE CERTEZA, VALIDACIÓN Y CUESTIONES JURÍDICAS

#### 8.1. Taxonomía de Niveles de Certeza Técnica
- **Certeza 1 (Hecho Fáctico Verificado en Producción):** Comprobado directamente en el binario `BDAS_Multiusuario.xlsm` o en la base de datos activa `Residencia_BE.accdb` mediante extracción, compilación o consulta de esquemas.
- **Certeza 2 (Solución Validada en Laboratorio):** Diseñada, implementada y demostrada mediante registros de pruebas unitarias verificables y reproducibles conservados en la documentación permanente del proyecto. *(Nota de rigor: aquellos componentes de laboratorio cuyos registros de prueba unitaria ejecutables no constan archivados en el repositorio, como clsDatabaseEngine y CapKey de la Fase 4, no pueden clasificarse en Certeza 2 y permanecen clasificados bajo Certeza 4)*.
- **Certeza 3 (Diseño Teórico o Propuesta Aprobada):** Solución formalmente acordada por la dirección arquitectónica (como las especificaciones de la Fase 6 / Rev. 5), cuya integración en el binario está pendiente.
- **Certeza 4 (Incertidumbre / Pendiente de Evidencia):** Cuestiones funcionales o históricas que no disponen de soporte documental concluyente.

#### 8.2. Cuestiones Jurídicas y Periodos de Retención
- **Parámetro Numérico de Retención en Código:**  
  El procedimiento `PurgarHistoricoRGPD` (`modDatabase.bas:4973-5003`) implementa la purga física sobre `Ordenes_Historico` (y `BDAS_Historico_Snapshot`) ejecutando la sentencia:
  ```vb
  sqlDelete = "DELETE FROM Ordenes_Historico WHERE FechaSalida <= DateAdd('yyyy', -" & anosRetencion & ", Date())"
  ```
  donde el valor numérico `anosRetencion` se define en la signatura con un valor predeterminado de **3 años** (`Optional ByVal anosRetencion As Long = 3`).
- **Deslinde de Conformidad Legal:**  
  1. *Naturaleza Exclusivamente Técnica del Procedimiento:* La ejecución de este procedimiento constituye únicamente un mecanismo de depuración de registros en la base de datos a petición de un administrador. La justificación regulatoria o idoneidad jurídica de dicho plazo de retención no emana de una garantía del código VBA, sino que **queda pendiente de validación jurídica independiente por parte de asesores legales especializados del proyecto**.
  2. *Delimitación de Alcance:* No debe afirmarse que la aplicación garantice por sí misma la conformidad integral con el Reglamento General de Protección de Datos (RGPD, UE 2016/679) ni con la LOPDGDD, la cual requiere medidas organizativas y normativas que exceden el software.

---

### 9. INCERTIDUMBRES TÉCNICAS QUE PERMANECEN ABIERTAS

1. **Trazabilidad Documental de Requisitos Fundacionales:**  
   No se dispone de las actas de especificación originales que dieron lugar a los requisitos `REQ-ORIG-01` a `REQ-ORIG-05`. Se conocen por su manifestación funcional as-built, catalogándose como "origen histórico pendiente de confirmar".
2. **Cronograma Formal de Inyección de la Arquitectura V5.2:**  
   La inyección de `clsDatabaseEngine`, la clave CapKey y la búsqueda de huecos ORD-10 en el binario de producción requiere una ventana de validación de campo multiusuario aún no programada.
3. **Tratamiento Definitivo de Archivos Huérfanos:**  
   Se requiere aprobación explícita de la dirección técnica para mover los 8 archivos huérfanos a un subdirectorio formal de archivo (`archive/`), evitando confusión en futuros desarrollos sin alterar el repositorio activo de forma imprevista.

---

### 10. REGLAS DE GOBERNANZA DEL DOCUMENTO MAESTRO Y REGISTRO DE VERSIONES

#### 10.1. Las Ocho Reglas de Mantenimiento Permanente
1. **Regla de Inmutabilidad As-Built:** Jamás se describirá como existente en producción una funcionalidad que no esté compilada en `BDAS_Multiusuario.xlsm`.
2. **Regla de Separación de Realidades:** Los Bloques I, II y III se mantendrán rigurosamente delimitados sin mezclar código real con aspiraciones de diseño.
3. **Regla de Trazabilidad Criptográfica:** Toda mención a binarios o bases de datos incluirá su hash SHA-256 verificado.
4. **Regla de Literalidad de Topónimos:** Se mantendrá permanentemente la escritura literal de `GIJÓN` o `GIJON` sin recurrir a secuencias numéricas o concatenaciones de caracteres.
5. **Regla de Cero Mojibake:** Todo cambio en la documentación o código se codificará en ANSI Windows-1252 o UTF-8 sin caracteres corruptos.
6. **Regla de Identificadores Estables:** Los códigos de defectos (`DEF-FIFO-01`, etc.) y de requisitos (`REQ-...`) no podrán renumerarse ni suprimirse; los resueltos se marcarán como cerrados.
7. **Regla de Validación Previa:** Ninguna modificación del Documento Maestro se publicará sin verificación previa de integridad y contraste contra el código fuente real.
8. **Regla de Registro de Versiones:** Cada actualización del documento incrementará su número de versión e incluirá una sección detallada de control de cambios.

#### 10.2. Registro de Versiones del Documento Maestro
| Versión | Fecha | Estado | Autor / Auditoría | Resumen de Cambios |
| :---: | :---: | :---: | :---: | :--- |
| **0.1** | 15/11/2026 | Superada | Antigravity / Auditoría Fase 6 | Consolidación documental inicial basada en las Órdenes 0, 0A, 0B y 0C. |
| **0.2** | 15/11/2026 | Histórica No Aceptada | Antigravity / Auditoría Orden 1A | Primera revisión de precisión técnica (detectadas discrepancias de consolidación). |
| **0.2.1**| 15/11/2026 | Histórica No Aceptada | Antigravity / Auditoría Orden 1B | Rectificación inicial de precisión técnica (rechazada por persistir errores de consolidación de la Orden 1A-B). |
| **0.2.2**| 15/11/2026 | Histórica No Aceptada | Antigravity / Cierre Orden 1C | Cierre técnico dirigido y extracción física de evidencias (rechazada por persistir imprecisiones en Certezas y Caso 900000). |
| **0.2.3**| 15/11/2026 | Histórica No Aceptada | Antigravity / Rectificación Orden 1D | Rectificación de Certezas y Caso 900000 (rechazada por justificación errónea de PK en `Ordenes_Historico`, ambigüedad en conexiones V5.2 y omisiones en mapeo por sede). |
| **0.2.4**| 15/11/2026 | Histórica No Aceptada | Antigravity / Rectificación Orden 1E | Rectificación de esquema en `Ordenes_Historico`, conexiones y Lease Lock (rechazada por persistir contradicciones en mapeo de SOTO, reintento automático en ORD-10, explicación errónea de bloqueo caducado y homologación no acreditada). |
| **0.2.5**| 16/11/2026 | Histórica No Aceptada | Antigravity / Rectificación Orden 1F | Rectificación de ORD-10, Lease Lock y SOTO (rechazada por persistir menciones residuales de homologación V5.2 no acreditada, equiparación de condiciones SQL de Lease Lock sin deslindar DateDiff < 0, referencias a rutinas de facturación/pago no acreditadas y discrepancia en metadatos de bytes). |
| **0.2.6**| 16/11/2026 | Histórica No Aceptada / Rectificada | Antigravity / Rectificación Orden 1G | Primera rectificación de seis bloques. La matriz y el análisis de sincronización por lotes quedaron incompletos al basarse en un extracto que no coincidía con el módulo de producción verificado en Orden 1H; se conserva como antecedente, no como versión vigente. |
| **0.2.7**| 16/11/2026 | **Candidata a Revisión Arquitectónica Independiente** | Dirección Arquitectónica / Rectificación Orden 1H | Identidad exacta del módulo de persistencia contra el binario activo; correspondencia confirmada de A:AB; persistencia por lote de NumFactura y EstadoPago; distinción entre transacción de lote y ausencia de Lease Lock por expediente; registro del defecto DEF-FIFO-07; actualización de la clasificación CapKey. |

---

### 11. CONTROL DE CAMBIOS DETALLADO (VERSIÓN 0.2 A VERSIÓN 0.2.7)

#### 11.1. Rectificaciones Incorporadas de Versión 0.2 a Versión 0.2.1
En cumplimiento de las instrucciones de la Orden 1B, la siguiente tabla detalla de forma exhaustiva las correcciones técnicas y rectificaciones incorporadas en la Versión 0.2.1 respecto a la versión 0.2:

| Ámbito / Sección | Contenido en Versión 0.2 | Rectificación Técnica en Versión 0.2.1 | Justificación y Evidencia Auditoría |
| :--- | :--- | :--- | :--- |
| **2.9. Hoja TURNOS-PRECIOS** | Mencionaba la hoja sin identificar su CodeName de forma canónica. | Documenta la correspondencia real: **CodeName `Hoja28`**, **Nombre visible `Turnos-Precios`**. | Comprobación en `vba-modules/Hoja28.cls` y `modDatabase.bas`. |
| **7.3. Recurso RibbonX** | Mencionaba genéricamente el archivo Ribbon sin precisar la ruta de paquete. | Documenta la ruta exacta en el contenedor OpenXML: **`customUI/customUI.xml`**, distinguiéndola del árbol de trabajo y descartando `customUI14.xml`. | Especificación de empaquetado OpenXML del libro XLSM. |
| **2.5 y 3.4. Token de Sesión** | Mantenía el formato obsoleto `Usuario_NombreEquipo_YYYYMMDD_HHMMSS`. | Sustituido íntegramente por el formato real: **`usuario@PC#HexGUID`** (API `CoCreateGuid`), documentando el acotamiento a <= 50 caracteres (longitud típica 46-47 chars). | Inspección de `modDatabase.bas:1029-1059` (`ObtenerTokenSesionInstancia`). |
| **2.4. Estructura de Columnas** | Afirmaba genéricamente 28 columnas idénticas para las tres residencias. | Tabla comparativa independiente detallando: GIJÓN (28 col, A:AB, ScrollArea A1:AD4000), SOTO (28 col, Col H Turno, ScrollArea A1:AD4000) y OVIEDO (33 col, Col S CamaSuple, Col T HabAsignadas, Col AB EstadoPago, ScrollArea A1:AG4000). | Inspección de `Hoja2.cls`, `Hoja20.cls`, `Hoja24.cls` y `ModuloSeguridadLocal.bas:683-741`. |
| **ANEXO A. Operaciones ADODB** | Breve mención resumida a 32 procedimientos. | Incorporación íntegra de la tabla exhaustiva de los **32 procedimientos ADODB** de la Orden 0B como Anexo A. | Requisito formal Orden 1B / Evidencia Orden 0B y 0C. |
| **2.3 y 7.1. Componentes VBA** | Desglose erróneo (116 .bas, 19 .frm, 11 .cls, 27 docs). | Corrección rigurosa: **106 módulos estándar `.bas`, 49 módulos de hoja, 1 módulo `ThisWorkbook`, 17 UserForms** (Total: 173 componentes). Diferenciación de 198 archivos en disco. | Conteo fáctico exacto de componentes en `BDAS_Multiusuario.xlsm`. |
| **6. Esquema Físico Access** | Nombres de campos inexactos en Usuarios (`Usuario`, `PasswordHash`), índices no acreditados en Bloqueos. | Corrección total: `Usuarios` con 8 campos (`Id`, `NombreUsuario`, `Clave`, `NombreCompleto`, `Rol`, `Residencias`, `Activo`, `FechaAlta`), `BloqueosHabitaciones` con `idx_Bloq_Res` en Residencia, `CodigosPostales` con 5 campos, `ContadorFacturas` con UK `(Residencia, Ejercicio)`, `BDAS_Historico_Snapshot` con 41 campos sin PK ni índices, y `EstadoHabitacionesCalendario` con 5 columnas físicas. | Catálogo ADOX auditado en Orden 0B. |
| **3.2 y 3.3. Firmas V5.2** | Atribuía `ObtenerMotorBaseDatos` a la clase y `ExecuteNonQuery` como `Long`. | Deslinde estricto: `clsDatabaseEngine` contiene `ExecuteNonQuery(...) As Boolean`, `ValidarCapacidadInterna` y `VincularInstanciaUnica`; `modDatabase.bas` contiene `m_EngineCapKey`, `ObtenerMotorBaseDatos` y `EjecutarDMLInterno`. | Inspección de fuentes de laboratorio Fase 4. |
| **2.7, 2.12, 4.1. Nombres Procedimientos** | Nombres conceptuales (`BuscarSiguienteNumeroOrdenLibre`, `ArchivarOrdenEnHistorico`, `VaciarHistoricoResidencia`). | Corrección a nombres reales del código: `ReservarSiguienteOrdenLibreBD` (`modDatabase:1403`), `ResolverNumOrdenProvisionalEnBD` (`SegLocal:1364`), `AsegurarOrdenExiste` (`modDatabase:1698`), `ArchivarOrdenesAntiguasBD` (`modDatabase:4784`), `PurgarHistoricoRGPD` (`modDatabase:4973`). | Trazabilidad de código en `modDatabase.bas` y `ModuloSeguridadLocal.bas`. |
| **2.7 y 4.1. Consulta MAX** | Consulta confusa con límite `< 900000`. | Deslinde de consultas: Reserva online usa `NumOrden < 900001` (`modDatabase:1430`); Reconciliación FIFO usa `NumOrden < MIN_ORDEN_PROVISIONAL` (900001) (`SegLocal:1373`). | Inspección de código en `modDatabase.bas` y `ModuloSeguridadLocal.bas`. |
| **2.7 y 5.2. Generación Provisional** | Afirmaba búsqueda en `SYS_COLA_SYNC`. | Precisa que `ObtenerSiguienteNumOrdenProvisional(wsResidencia)` examina la hoja de residencia recibida como argumento, no `SYS_COLA_SYNC`. | Inspección de `ModuloSeguridadLocal.bas:598-615`. |
| **2.11, 3.6, 5.3. Tratamiento de `900000`** | Describía `900000` como "guardia o zona neutra". | Tratamiento puramente neutral: valor frontera excluido de rangos oficial y provisional, sin atributos adicionales. | Directriz expresa Orden 1B. |
| **3.5. Búsqueda de Huecos** | Afirmaba aborto ante celdas vacías intermedias. | Precisa que se permiten celdas vacías intermedias, validando todos los valores no vacíos y continuando desde `valorUltimaFila + 1` si no hay huecos. | Directriz consensuada Orden 1B. |
| **2.12 y 8.2. Purga Histórico** | Usaba `Date - 1095`. | Refleja la expresión real: `DateAdd('yyyy', -anosRetencion, Date())` en `PurgarHistoricoRGPD`. | Inspección de `modDatabase.bas:4986`. |
| **7.2. Archivos Huérfanos** | Lista no normalizada. | Relación acreditada de los 8 archivos: `CalcularFacturaEstudiante.bas`, `modMigracion.bas`, `ModuloAjusteZoomHojas.bas`, `ModuloCrearHojaConfig.bas`, `ModuloGestionCacheCalendarios.bas`, `SolapesAceptadosListado.bas`, `TemporalEliminarLuego.bas`, `Hoja34_TarifasUnidad.cls`. | Cotejo `vba-modules/` vs `BDAS_Multiusuario.xlsm`. |
| **4.1. Referencias de Líneas** | Líneas fuera de rango en `ModuloSeguridadLocal.bas` (> 1648). | Corregidas todas las referencias contra los rangos reales de la Orden 0C (`ModuloSeguridadLocal.bas` tiene exactamente 1.648 líneas). | Verificación física de longitudes de archivo. |
| **5.4. Propuestas Externas** | Incluía SQL Server, PostgreSQL, Web SPA, WebSockets. | Eliminadas de REQ-PROP por ser ideas externas no autorizadas fuera del alcance arquitectónico del proyecto. | Directriz expresa Orden 1B. |
| **Metadatos y Estado** | Figuraba como "Vigente Auditada". | Figurará como **`CANDIDATA A REVISIÓN ARQUITECTÓNICA INDEPENDIENTE`**. | Directriz formal de gobierno documental. |

#### 11.2. Rectificaciones Incorporadas de Versión 0.2.1 a Versión 0.2.2 (Orden 1C)
A continuación se auditan punto por punto las rectificaciones empíricas incorporadas en v0.2.2 tras la ejecución del Cierre Técnico Dirigido de la Orden 1C:

| Num | Ámbito / Apartado | Estado en v0.2.1 (Defecto / Imprecisión) | Rectificación Aplicada en v0.2.2 | Base de Evidencia Fáctica (Orden 1C) |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Token de Sesión (Sección 2.5)** | Estimación vaga de "46-47 caracteres" y omisión de la fórmula exacta. | Desglose aritmético exacto (untruncado: $\text{Len}(usr) + \text{Len}(pc) + 14$; truncado fijo: 50 caracteres exactos; ejemplo estándar: 42 caracteres). Aclarada la omisión de `Data3` y `Data4` de `GUID_TYPE` (solo 12 hex chars usados). | `modDatabase.bas:1029-1059`. |
| **2** | **ScrollArea SOTO (Sección 2.4)** | Mención no justificada a rango `A1:AE4000` sin aclarar el origen de `A1:AC4000`. | Fijada la ScrollArea canónica en `A1:AD4000` (`Hoja20.cls:57`) y documentado que la confusión de `A1:AC4000` procede del filtro de eventos de columna 29 en `Hoja20.cls:36`. | Inspección global de código en `vba-modules\Hoja20.cls`. |
| **3** | **Esquema Access (Sección 6)** | Inclusión de 6 tablas especulativas inexistentes; error en `EstadoHabitacionesCalendario` (5 cols, con `Habitacion`) y `ListaNegra` (índice en `Nombre`). | Refutación de las 6 tablas inexistentes. Fijadas las 11 tablas reales verificadas por ADOX: `EstadoHabitacionesCalendario` posee 4 cols (sin `Habitacion`, índice único en Residencia/Fila) y `ListaNegra` posee índice único solo en `DNI`. | Inspección OLEDB/ADOX directa sobre `Residencia_BE.accdb`. |
| **4** | **Ciclo de Vida Histórico (Sección 2.12)** | Descripción genérica sin detalle de filtros ni transaccionalidad. | Detallada la condición obligatoria de datos reales (`DNI` o `Nombre` no vacíos), borrado previo de duplicados en histórico, transacción atómica `BeginTrans`/`CommitTrans`, vaciado exclusivo de sede activa y 2 fases de archivado. | `modDatabase.bas:4261-5003`. |
| **5** | **Firmas V5.2 (Sección 3.2)** | Presencia de elipsis `(...)` y tipado ambiguo en `clsDatabaseEngine`. | Firmas completas literales sin elipsis en tabla dedicada; fijado el retorno `Boolean` en `ExecuteNonQuery` y `EjecutarUpdateCondicionalOrden`. | Documentación técnica consolidada Fase 4. |
| **6** | **Ciclo CapKey (Sección 3.3)** | Descripción incompleta de la transmisión y validación oracular. | Documentado el flujo completo: generación por CSPRNG (`GenerarHexCSPRNG32Bytes`, 256 bits), enlace por `VincularInstanciaUnica` y validación en tiempo constante por XOR acumulado (`CompararCadenasTiempoConstante`). | Código de laboratorio y dictámenes Fases 3b/4. |
| **7** | **Escenarios FIFO (Sección 3.7)** | Flujo FIFO descrito a alto nivel sin diferenciar los tres escenarios de concurrencia. | Trazabilidad exhaustiva de los tres escenarios reales: Escenario A (bloqueo activo, rollback y refresco), Escenario B (bloqueo caducado, takeover) y Escenario C (alta previa vía `AsegurarOrdenExiste`). | `modDatabase.bas:1737-1993` y `ModuloSeguridadLocal.bas`. |
| **8** | **Caso Límite 900000 (Sección 2.7 y 4.1)** | Tratamiento superficial de 900000 y diagnóstico erróneo de colisión en `DEF-NUM-01`. | Trazabilidad completa de 900000 (generable por `MAX+1` y aceptado en entrada manual online al dar `EsNumOrdenProvisional=False`). Reclasificado `DEF-NUM-01` como falta de búsqueda de huecos y ausencia de techo en 899999. | `modDatabase.bas:1403-1472`, `1527-1696` y `ModuloSeguridadLocal.bas:617`. |
| **9** | **Requisitos de Negocio (Sección 5 y 8.2)** | Afirmación de "sin huecos garantizado" en producción y falta de advertencias sobre RD 933/2021. | Adecuación al principio de tres realidades: producción usa `MAX+1` (sin huecos es objetivo); advertencia expresa de ausencia de remisión telemática automatizada a plataformas externas. | `modDatabase.bas:1403` y análisis normativo RD 933/2021. |
| **10** | **Conciliación ADODB (Anexo A)** | Omisión de procedimientos orquestadores compuestos de persistencia. | Reconciliación matemática: 32 procedimientos directos DML/DDL + 2 orquestadores compuestos (`ArchivarOrdenesAntiguasBD` y `GestionarEntradaNumOrdenHoja`) = 34 procedimientos totales. | Escaneo sintáctico completo de los 134 procedimientos en `modDatabase.bas`. |

---


#### 11.3. Rectificaciones Incorporadas de Versión 0.2.2 a Versión 0.2.3 (Orden 1D)
La presente versión 0.2.3 incorpora de forma verificable las rectificaciones ordenadas en la Orden 1D tras la extracción física de evidencias del binario candidato de laboratorio Fase 4 y la auditoría cruzada con el backend de producción:

| Nº | Ámbito / Apartado | Deficiencia Subsanada de v0.2.2 | Rectificación Incorporada en v0.2.3 | Evidencia Técnica Primaria |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Apartado 1 (Tres Realidades)** | Mezcla en Bloque II de componentes de laboratorio y propuestas de diseño. | Delimitación neta: 1) Código en laboratorio: `clsDatabaseEngine` V5.2 y `CapKey` V5.2; 2) Diseño aprobado pendiente de implementación: ORD-10 Rev. 5, FIFO Rev. 5, Rangos canónicos y TAR-10. | Binario `BDAS_Multiusuario_Fase4.xlsm` y código As-Built. |
| **2** | **Apartados 2.11 y 2.7 (Caso 900000)** | En 2.11 se afirmaba erróneamente que 900000 estaba excluido en producción. | Erradicación de la afirmación de exclusión en producción. Distinción explícita: As-Built actual lo genera/acepta online; exclusión es un requisito estricto de la arquitectura objetivo. | `modDatabase.bas:1403-1696` y `ModuloSeguridadLocal.bas:617`. |
| **3** | **Apartados 3.8 y 3.9 (Certezas)** | TAR-10 figuraba como Certeza 2 en Versión Referencia; diseño con estados de laboratorio. | TAR-10 reclasificado a **Certeza 3 (Diseño aprobado no implementado)**, aclarando que Versión Referencia solo posee persistencia base y optimización ScreenUpdating/Calculation. ORD-10, FIFO Rev. 5 y Rangos en Certeza 3; `clsDatabaseEngine` y CapKey en Certeza 2 de laboratorio; Token de sesión en Certeza 1 con limitaciones. | Binarios de Versión Referencia y Laboratorio Fase 4. |
| **4** | **Apartado 4.1 (Fichas DEF)** | Imprecisiones conceptuales en `DEF-FIFO-02..05`, `DEF-NUM-03`, `DEF-ORD-01`, `DEF-HIST-01..02` y `DEF-ARCH-01`. | - `DEF-FIFO-02/03`: Foco en ausencia de verificación de identidad de expediente (sin cotejo DNI/Id).<br>- `DEF-FIFO-04`: Limitado al rechazo de Lease Lock con refresco y sobreescritura destructiva local.<br>- `DEF-FIFO-05`: Ausencia de transacción compartida entre inserción de esqueleto y datos.<br>- `DEF-NUM-03`: Conversiones atribuidas con líneas exactas (`modDatabase.bas:1548`, etc.).<br>- `DEF-ORD-01`: Se destaca que Access previene duplicados físicos; defecto real es falta de reutilización de huecos y dependencia de MAX+1.<br>- `DEF-HIST-01`: Transacción atómica existe; riesgo real es sustitución de histórico previo.<br>- `DEF-HIST-02`: Eliminadas menciones a 2FA; documentado Caso B real.<br>- `DEF-ARCH-01`: Separación neta entre Access y borrado visual en Excel. | Inspección línea a línea en `modDatabase.bas` y `ModuloSeguridadLocal.bas`. |
| **5** | **Apartado 2.12 (Ciclo de Vida)** | Vaciado de sede describía borrado total; no se alertaba de sustitución de históricos. | Se documenta el Caso B real en `EliminarTodasOrdenesSedeBD` (solo borra históricos coincidentes) y se documenta el riesgo de sustitución en `EliminarOrdenBD` y `EliminarOrdenesLoteBD`. | `modDatabase.bas:4261-4659`. |
| **6** | **Apartados 5.2 y 5.3 (Requisitos)** | `REQ-DEV-05` figuraba como purga periódica; `REQ-OBJ-05` figuraba en Versión Referencia. | `REQ-DEV-05` se define como "Purga de histórico invocable" (desde `frmPanelRGPD`). `REQ-OBJ-05` se fija en Certeza 3 "Aprobado en Diseño" documental. | `frmPanelRGPD.frm` y auditoría de la Versión Referencia. |
| **7** | **Apartado 8.2 (Cuestiones Jurídicas)** | Se afirmaba mandato legal de 3 años y remisión telemática externa. | Se sustituye por descripción técnica objetiva: `anosRetencion = 3` es un parámetro por defecto sujeto a validación jurídica externa. Eliminada toda mención a plataformas externas de hospedajes. | `modDatabase.bas:4973`. |
| **8** | **Anexo A (Inventario ADODB)** | Fila 14 figuraba como `DELETE FROM`. Texto introductorio excluía orquestadores de forma ambigua. | Fila 14 corregida a `UPDATE (Activo=False)` (`modDatabase.bas:3158-3166`). Líneas de `AsegurarEsquemaConcurrenciaBD` fijadas en 1118-1200. Texto clarifica 32 directos DML/DDL y 2 orquestadores compuestos de persistencia. | Auditoría sintáctica de `modDatabase.bas`. |
| **9** | **Apartado 6 (Modelo Físico Access)**| Falta de trazabilidad expresa catálogo-consumidores en `ListaNegra` y `LogActividad`. | Vinculación explícita de campos físicos con `AnadirAListaNegra`, `RetirarDeListaNegra` e `InsertarLog`. | Consultas ADOX y código de `modDatabase.bas:3114-3201`. |
| **10** | **Apartado 2.5 (Token de Sesión)** | No se documentaba la extracción parcial de 12 hex ni la debilidad del fallback. | Se documenta que solo se leen `Data1` y `Data2` (12 hex, 48 bits) y que el fallback `Timer`/`Rnd` carece de unicidad criptográfica demostrada. | `modDatabase.bas:1029-1059`. |

#### 11.4. Rectificaciones Incorporadas de Versión 0.2.3 a Versión 0.2.4 (Orden 1E)
La versión 0.2.4 incorpora las correcciones técnicas obligatorias ordenadas tras la revisión independiente de la versión 0.2.3, contrastadas directamente con el código fuente de producción, el binario de laboratorio Fase 4 y el catálogo ADOX de Access:

| Nº | Ámbito / Apartado | Deficiencia Subsanada de v0.2.3 | Rectificación Incorporada en v0.2.4 | Evidencia Técnica Primaria |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Apartados 2.12, 4.1 (`DEF-HIST-01/02`)** | Afirmaba que el borrado previo en `Ordenes_Historico` prevenía colisión de PK / error 3022. | Erradicación total de la justificación de PK/UK. `Ordenes_Historico` carece de PK e índice tiene `Unique=False`. El borrado previo es una política procedimental de sustitución. Se documenta la pérdida destructiva de históricos legítimos si se reasignan números. | Inspección física ADOX/OLEDB y `modDatabase.bas:4261-4659`. |
| **2** | **Apartado 2.12 (`EliminarTodasOrdenesSedeBD`)** | Ambigüedad sobre el alcance del borrado en vaciado de sede. | Reafirmado estrictamente el Caso B: solo elimina históricos coincidentes con órdenes activas con datos reales; el resto del histórico de la sede permanece intacto. | `modDatabase.bas:4556-4659`. |
| **3** | **Apartado 3.2 (Conexiones V5.2)** | Afirmaba que `clsDatabaseEngine` administra una única conexión ADODB activa persistente. | Rectificado: la instancia del motor es singleton, pero abre y cierra conexiones ADODB puntuales por cada operación (`cn.Open` / `cn.Close`). | `clsDatabaseEngine.cls:233-239` y `411-450`. |
| **4** | **Apartado 3.2 (Firmas V5.2)** | Signatura de fachada `EjecutarUpdateCondicionalOrden` figuraba con ámbito `Public / Friend`. | Corregido a ámbito estrictamente `Private Function` en `modDatabase.bas:8355-8395`. | Extracción binaria de `modDatabase.bas` en Fase 4. |
| **5** | **Apartado 3.4 (SQL Lease Lock V5.2)** | Presentaba sentencia esquemática ficticia con `WHERE Id = [IdExpediente]`. | Sustituido por el SQL real extraído de `clsDatabaseEngine.cls` identificando por `(NumOrden, Residencia)` y validando Lease Lock de 120 s. | `clsDatabaseEngine.cls:365-433`. |
| **6** | **Apartado 2.4 (Mapeo por Sede)** | Se describía genéricamente como «28 campos de la sede». | Tabla y desglose en 3 bloques: GIJÓN (28 col leídas, 28 escritas, col 28 no persistida), SOTO (28 leídas, 26 escritas, col 8 Turno, cols 17-18 y 28 no persistidas), OVIEDO (28 leídas [cols 29-33 no leídas], 29 escritas, col 19 CamaSuple, col 28 EstadoPago). | `modDatabase.bas:1737-1993` y controladores de hoja. |
| **7** | **Apartados 3.7 y 4.1 (`DEF-FIFO-05`)** | Describía `INSERT INTO Ordenes (NumOrden, Residencia)` simplificado. | Reproducida la sentencia literal de 8 campos insertados por `AsegurarOrdenExiste` y su ejecución independiente fuera de la transacción posterior. | `modDatabase.bas:1698-1732`. |
| **8** | **Apartado 4.1 (`DEF-FIFO-04`)** | Calificaba la pérdida de cambios como «destrucción silenciosa». | Rectificado: distingue la alerta modal explícita (`MsgBox vbExclamation`) de la pérdida irreversible de datos por carecer de mecanismo de recuperación/stash local. | `modDatabase.bas:1834-1839`. |
| **9** | **Apartado 2.5 (Token de Sesión)** | Presentación susceptible de interpretarse como identificador criptográfico garantizado. | Homogeneizado: 48 bits (Data1 y Data2), fallback débil `Timer`/`Rnd`, diseñado para diferenciar instancias sin garantía criptográfica de unicidad absoluta. | `modDatabase.bas:1029-1059`. |
| **10** | **Apartados 4.1 y 5.4 (Propuestas)** | Propuestas figuraban como requisitos o con estados formales («En Estudio»). | Renombradas a `PROP-ARQ-01` y `PROP-ARQ-02` ("Propuesta No Aprobada"). En fichas DEF se deslindan decisiones aprobadas (ORD-10, V5.2) de alternativas propuestas. | Criterio formal Orden 1E. |
| **11** | **Apartado 7.2 (Archivos Huérfanos)** | Atribuía causas de obsolescencia o carácter experimental sin acreditación documental. | Catálogo factual estricto de los 8 archivos ausentes en el binario compilado sin explicaciones especulativas. | Inventario físico del repositorio. |
| **12** | **Anexo A (Introducción Duplicada)** | Contenía dos párrafos introductorios consecutivos redundantes. | Eliminado el segundo párrafo redundante. | Revisión editorial Anexo A. |

#### 11.5. Rectificaciones Incorporadas de Versión 0.2.4 a Versión 0.2.5 (Orden 1F)
La versión 0.2.5 incorpora las rectificaciones obligatorias establecidas por la dirección arquitectónica en la Orden 1F, contrastadas documentalmente contra el código fuente, la estructura física de las hojas Excel y el esquema de Access:

| Nº | Ámbito / Apartado | Deficiencia Subsanada de v0.2.4 | Rectificación Incorporada en v0.2.5 | Evidencia Técnica Primaria |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Apartado 3.5 (ORD-10, Punto 8)** | Permitía el reintento y asignación automática con el siguiente número libre ante colisión. | Prohibición absoluta de reasignación automática: el candidato se presenta al operador; ante colisión concurrente, la reserva fracasa, se informa al usuario, se recalcula un nuevo candidato y se exige confirmación expresa antes de intentar reservarlo. | Decisión funcional aprobada / Especificación Orden 1F. |
| **2** | **Apartado 3.4 (Lease Lock)** | Afirmaba que un bloqueo caducado produce cero registros afectados (`affected = 0`). | Deslinde neto: Caso A (bloqueo vigente de otro puesto produce 0 registros y aborta); Caso B (bloqueo caducado $\ge 120$ s produce 1 registro afectado y transfiere la titularidad). Armonizado con Escenario B de 3.7. | Código literal en `clsDatabaseEngine.cls:365-433` y `modDatabase.bas:1791-1800`. |
| **3** | **Apartado 2.4 (Mapeo SOTO)** | Contradicción interna: identificaba Cols 17-18 como `NumHabIndividuales`/`NumHabDobles` y como `Observaciones`/`NumHab`. | Comprobación física en `sheet23.xml`: Col 17 es `PRECIO APTO.`, Col 18 es `SUPLE. OCUPAN.` y Col 19 es `APARTAMENTO`. `SincronizarFilaHaciaAccess` lee Cols 17-18 pero no las asigna a Access. | `BDAS_Multiusuario.xlsm:sheet23.xml` y `modDatabase.bas:1938-1948`. |
| **4** | **Apartado 2.4 (Cinco Dimensiones)** | Mapeos dispersos sin separación estricta entre lectura en memoria y campos en Access. | Tabla canónica unificada con cinco dimensiones (columna física, encabezado real, lectura en sincronización, asignación en Access y persistencia por otras rutinas), acreditando columnas auxiliares y ScrollArea. | Inspección de hojas maestras, `modDatabase.bas` y `ModuloSeguridadLocal.bas:683-741`. |
| **5** | **Apartado 2.5 (Token de Sesión)** | Conservaba afirmaciones de unicidad garantizada por instancia pese a tener solo 48 bits y fallback débil. | Erradicación total de afirmaciones de unicidad garantizada. Se explicita su finalidad operativa de diferenciación en ejecución normal, documentando su debilidad pseudo-aleatoria ante fallos de API. | Código literal en `modDatabase.bas:1029-1059`. |
| **6** | **Apartados 3.2 y 3.9 (Homologación V5.2)** | Calificaba `clsDatabaseEngine` y CapKey como «Implementado y Homologado» sin registros de prueba en repositorio. | Deslinde estricto: código presente físicamente en binario de laboratorio Fase 4, pero homologación clasificada en Certeza 4 (Incertidumbre / Pendiente de Evidencia de Pruebas) por ausencia de registros reproducibles en repositorio. | Inventario técnico del repositorio y taxonomía de certezas 8.1. |
| **7** | **Inmutabilidad Global** | Verificación parcial limitada a front-end y back-end de producción en el informe anterior. | Comprobación criptográfica completa individualizada de los 7 elementos protegidos del proyecto (producción, referencia, Golden Master y Fase 4), con hashes de 64 caracteres hex. | Verificación SHA-256 en sistema de archivos. |

---



#### 11.6. Rectificaciones Incorporadas de Versión 0.2.5 a Versión 0.2.6 (Orden 1G)
La versión 0.2.6 incorpora las rectificaciones obligatorias establecidas por la dirección arquitectónica en la Orden 1G, eliminando discrepancias residuales y homogeneizando la documentación contra el código vigente:

| Nº | Ámbito / Apartado | Deficiencia Subsanada de v0.2.5 | Rectificación Incorporada en v0.2.6 | Evidencia Técnica Primaria |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Apartados 1, 3.2, 3.9, 8.1 (Homologación V5.2)** | Presencia de afirmaciones residuales («lo validado en laboratorio es su ejecución encapsulada», «componentes implementados y homologados en laboratorio»). | Erradicación total de afirmaciones residuales de validación u homologación acreditada. Formulación conceptual uniforme: código físicamente integrado en `BDAS_Multiusuario_Fase4.xlsm` y extraído, pero homologación no acreditada con registros reproducibles en el repositorio permanente, manteniéndose clasificados en Certeza 4. | `BDAS_Multiusuario_Fase4.xlsm` y archivo permanente de pruebas. |
| **2** | **Apartados 3.4 y 3.7 (Condiciones SQL Lease Lock)** | Presentación como equivalentes de sentencias SQL no idénticas entre producción y laboratorio V5.2. | Deslinde riguroso: Producción (`modDatabase.bas:1791-1800`) incorpora `DateDiff < 0` para marcas futuras por desincronización de reloj y usuario de fallback; Laboratorio (`clsDatabaseEngine.cls:375-383`) no contiene `DateDiff < 0`. Documentación de los cuatro escenarios operativos (bloqueo vigente, caducado >= 120s, marca futura y registro inexistente). | Código fuente de `clsDatabaseEngine.cls:375-383` y `modDatabase.bas:1791-1800`. |
| **3** | **Apartado 2.4 (Facturación y Estado de Pago)** | Citaba rutinas inexistentes `ActualizarFacturaBD` (`:2088-2129`) y `ActualizarEstadoPagoBD` (`:2133-2174`) que colisionaban con `SincronizarLoteImportacionBD`. | Verificación en código fuente: dichas rutinas no existen en el código del proyecto. Para los campos de facturación y pago, la 5ª columna se corrige con la calificación canónica exacta: «No se ha acreditado una rutina de persistencia independiente para este valor en el alcance de código examinado». | Búsqueda exhaustiva en `vba-modules/` y código de `modDatabase.bas`. |
| **4** | **Apartado 2.4 (Matriz de Cinco Dimensiones)** | La 5ª dimensión contenía descripciones de eventos, navegación y notas visuales que no correspondían semánticamente a persistencia. | Denominación corregida de la 5ª columna a «Persistencia en Access mediante otra rutina independiente (si está acreditada)». Contenido restringido estrictamente a las rutinas de reserva atómica de `NumOrden` (Col. 1) y a la calificación formal de no acreditada en el resto. Información de navegación (`ScrollArea`) y eventos trasladada a notas explicativas. | Análisis semántico y estructural de las hojas operativas. |
| **5** | **Metadatos y Tamaños (Bytes vs Caracteres)** | Discrepancia en Informe 1F entre tamaño en bytes (158.703) y caracteres (154.385). | Unificación estricta y deslinde formal: tamaño físico en bytes (`os.path.getsize`) diferenciado de número de caracteres, calculado sobre el archivo final y mantenido idéntico en todas las menciones. | Medición sobre sistema de archivos Windows NTFS. |
| **6** | **Scripts Auxiliares y Estado de Git** | Informe 1F no detallaba la ubicación de scripts auxiliares ni aportaba comparación verificable del estado de Git anterior a la orden. | Relación exhaustiva de los scripts auxiliares de la Orden 1F (`build_v025.py`, `build_report_1f.py`, `verify_orden1f.py`, etc.) en el directorio `scratch` fuera del repositorio. Declaración formal de «NO ACREDITADO» para la comparación de Git de la Orden 1F por falta de captura inicial verificable. | Registros de ejecución y comandos de sistema. |


---

**Rectificación posterior:** las afirmaciones de esta tabla describen lo que se documentó en v0.2.6 y se conservan como historial. Los apartados 2.4, 3.4 y Anexo A quedan sustituidos por la rectificación 11.7 de v0.2.7. En particular, la ausencia de procedimientos denominados `ActualizarFacturaBD` y `ActualizarEstadoPagoBD` no significa ausencia de persistencia de `NumFactura` o `EstadoPago`, porque ambos campos se escriben mediante `SincronizarLoteImportacionBD`.

#### 11.7. Rectificaciones Incorporadas de Versión 0.2.6 a Versión 0.2.7 (Orden 1H)
La versión 0.2.7 se ha rectificado directamente a partir de la extracción del módulo `VBA/modDatabase` del binario activo y de la comparación byte a byte con `vba-modules/modDatabase.bas`. No se ha delegado la redacción del Documento Maestro en la IA ejecutora.

| Nº | Ámbito / Apartado | Deficiencia o incertidumbre detectada en v0.2.6 | Rectificación incorporada en v0.2.7 | Evidencia técnica primaria |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **Identidad del módulo** | La correspondencia entre el código fuente inspeccionado y el módulo integrado no estaba cerrada mediante hash y cotejo byte a byte en la documentación recibida. | Se documenta identidad exacta: 298.794 bytes, 7.231 líneas y SHA-256 `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` tanto en repositorio como en el módulo extraído del binario activo. | Informe de Verificación Final de Identidad del Módulo de Persistencia, 16/11/2026, apartados 1.1–1.5. |
| **2** | **Apartado 2.4: matriz de persistencia** | La quinta dimensión no reconocía a `SincronizarLoteImportacionBD` como una segunda ruta de persistencia para los campos que procesa. | Se identifica expresamente la persistencia por lotes de los campos mapeados y se mantiene el detalle diferenciado por residencia. `NumFactura` se asigna desde C; `EstadoPago` desde AA en GIJÓN/SOTO y AB en OVIEDO. | Módulo integrado, líneas 2021–2238, y encabezados A:AC extraídos de las tres hojas. |
| **3** | **Persistencia de factura y pago** | La conclusión de 1G confundía la inexistencia de procedimientos con nombres específicos (`ActualizarFacturaBD`, `ActualizarEstadoPagoBD`) con la inexistencia de persistencia de esos campos. | Se conserva la constatación de que esos dos nombres de procedimiento no están acreditados, pero se rectifica el fondo: `SincronizarFilaHaciaAccess` y `SincronizarLoteImportacionBD` asignan `NumFactura` y `EstadoPago`, con preservación de valores existentes cuando la celda correspondiente llega vacía en registros ya existentes. | Módulo integrado, líneas 1898–1903, 1924–1967, 2128–2133 y 2191–2196. |
| **4** | **Transacción de importación masiva** | La documentación derivada del extracto anterior indicaba que la rutina no abría una transacción y usaba `UpdateBatch`. | Se rectifica: el módulo integrado ejecuta `cn.BeginTrans`, aplica `rs.Update` por registro, confirma con `cn.CommitTrans` al final y revierte con `cn.RollbackTrans` ante error. No usa `rs.UpdateBatch` en esta versión integrada. | Módulo integrado, líneas 2074–2085, 2198–2218 y 2223–2238. |
| **5** | **Concurrencia de la importación masiva** | El Anexo A atribuía Lease Lock a `SincronizarLoteImportacionBD`. | Se elimina esa atribución: la rutina por lotes no realiza adquisición condicional ni verifica la titularidad/antigüedad del Lease Lock por expediente; escribe el token actual y actualiza registros existentes. Se añade `DEF-FIFO-07`. | Módulo integrado, líneas 2021–2238; comparación con el bloque condicional de `SincronizarFilaHaciaAccess`, líneas 1789–1803. |
| **6** | **Apartado 3.4: Lease Lock de producción** | Los escenarios de bloqueo vigente no explicitaban todas las precondiciones necesarias para que el `UPDATE` afectara cero filas y no explicaban el tratamiento funcional cuando no existe registro. | Se mantienen las cláusulas SQL reales, se especifican las condiciones simultáneas del bloqueo vigente y se distingue el resultado `recsLock = 0` por bloqueo del resultado por clave inexistente. Se documenta el efecto de `permitirAlta`. | Módulo integrado, líneas 1791–1803 y 1807–1863. |
| **7** | **REQ-OBJ-03 / CapKey** | El requisito de mediación CapKey aparecía como «Aprobado en Laboratorio», lo que podía confundirse con homologación funcional acreditada. | Se clasifica como diseño aprobado, código presente en laboratorio y homologación funcional no acreditada (Certeza 4), coherente con el apartado 8.1. | Evidencias de Fase 4 y criterio de certeza vigente. |
| **8** | **Identificación de hojas / CodeName** | El informe de extracción de la Orden 1H confundió identificadores de partes XML (`sheetN.xml`) con CodeName VBA y dejó la correspondencia pendiente. | Estado inicial de la Orden 1H: pendiente. La Orden 1J lo resuelve con el volcado estructural del libro activo: `Hoja2` = RESIDENCIA GIJÓN, `Hoja20` = RESIDENCIA SOTO y `Hoja24` = RESIDENCIA OVIEDO. No se renombra ni reasigna ningún módulo. | Volcado estructural JSON del libro activo aportado para la revisión 1J, cruzado con la relación de hojas y el inventario previo de módulos de hoja. |
| **9** | **Alcance de esta rectificación** | El extracto técnico anterior de las 17:19 describía firmas y comportamiento distintos en estas dos rutinas (parámetro `resTarget`, `UpdateBatch` y flujo sin transacción explícita). | Se considera superado para describir la producción por el cotejo posterior con el módulo del binario activo. La nueva evidencia no autoriza a extrapolar esa corrección a procedimientos ajenos a las dos rutinas transcritas en Orden 1H, que no se reauditan en esta revisión. | Comparación del contenido técnico aportado a las 17:19 con la extracción exacta del módulo activo de las 19:26. |
| **10** | **Git / Orden 1F** | El Criterio 8 de Orden 1G seguía sin evidencia inicial de Git previa a Orden 1F. | Se conserva **NO ACREDITADO**. La Orden 1H no reconstruye ni presume el estado histórico que no quedó capturado. | Informe 1G y registros de Git disponibles. |

**Alcance y cautela:** esta revisión actualiza el mapa de persistencia y las dos rutinas transcritas en la evidencia de Orden 1H. La coincidencia byte a byte certifica que el archivo fuente corresponde al módulo integrado, pero no convierte en verificados por esta orden todos los demás procedimientos de `modDatabase.bas`. Las conclusiones anteriores sobre procedimientos que no se han vuelto a transcribir se conservan como antecedentes y no se presentan como revalidados en Orden 1H.

#### 11.8. Precisiones Complementarias a la Versión 0.2.7 (Orden 1J)
La revisión complementaria 1J resuelve la identificación de los CodeName de las hojas maestras y precisa el alcance del literal `PENDIENTE` en las rutas de alta. No modifica el análisis de concurrencia, el catálogo de defectos, el índice ni la estructura o el recuento del Anexo A; solo actualiza la identificación CodeName de sus llamadores.

| Nº | Ámbito / Apartado | Estado anterior | Precisión incorporada | Evidencia técnica |
| :---: | :--- | :--- | :--- | :--- |
| **1** | **CodeName de hojas maestras** | La Orden 1H dejó pendiente la correspondencia entre los nombres `sheetN.xml` y los CodeName VBA. | Se confirma `Hoja2` = RESIDENCIA GIJÓN (`sheet2.xml`), `Hoja20` = RESIDENCIA SOTO (`sheet23.xml`) y `Hoja24` = RESIDENCIA OVIEDO (`sheet34.xml`). Se aclara que los nombres de partes XML no son CodeName. También se identifican `Hoja1` como Calendario GIJÓN, `Hoja23` como Calendario OVIEDO y `Hoja34` como TARIFAS UNIDAD. | Volcado estructural JSON del libro activo y cotejo con el inventario previo de módulos de hoja. |
| **2** | **`EstadoPago` / valor inicial `PENDIENTE`** | La sección 2.4 describía correctamente que la rutina por lotes asigna el contenido de la celda, pero podía interpretarse como ausencia de cualquier inicialización de `PENDIENTE` en el sistema. | Se aclara que `IntentarReservarOrdenBD`, en la rama de inserción de un número libre, establece explícitamente `EstadoPago = 'PENDIENTE'`. Esto no acredita un valor predeterminado general configurado en el campo Access ni cambia el comportamiento de la importación por lotes. | `modDatabase.bas`, `IntentarReservarOrdenBD`, líneas 1358–1369, cotejado con el módulo integrado. |

**Control de alcance:** la Orden 1J es una precisión documental. No modifica código, no declara homologación funcional de V5.2/CapKey, no eleva la certeza de las pruebas y no altera la condición de la v0.2.7 como candidata a revisión arquitectónica independiente.

---

#### 11.9. Precisión Complementaria — Orden 1K (Límite Real de Reintentos en Reserva Online)

La revisión independiente de la v0.2.7 detectó que el apartado 2.7 y la ficha `DEF-ORD-01` describían un máximo de 20 iteraciones para `ReservarSiguienteOrdenLibreBD`. Esa cifra no coincide con el código As-Built integrado en producción.

- El código establece `maxIntentos = 5` (`modDatabase.bas:1425`).
- El bucle `Do While Not reservado And intentos < maxIntentos` se ejecuta como máximo cinco veces (`modDatabase.bas:1429-1467`).
- En cada iteración vuelve a calcular `ObtenerUltimoNumOrdenBD(resTarget) + 1`; no busca huecos interiores.

Se rectifican exclusivamente las dos menciones al límite de 20 intentos, sustituyéndolas por el límite real de cinco. No se modifica el diagnóstico de dependencia de `MAX+1`, la ausencia de búsqueda de huecos, la vulnerabilidad del valor 900000 ni el diseño objetivo ORD-10. Esta precisión es documental y no modifica el código ni el catálogo de 18 defectos.

#### 11.10. Conciliación de `ConfigTarifasCeldas` — Orden 1K

La consulta de solo lectura realizada sobre la base Access de producción el 16/11/2026 concilia los **291 registros activos** con los rangos realmente devueltos por `ObtenerRangosEditablesTarifas`:

- `TARIFAS UNIDAD`: `A3:C7` (15), `E3:N9` (70) y `E13:N19` (70) = **155 registros**. Las 42 coordenadas `K3:N9` y `M13:N19` que habían quedado fuera de una comprobación anterior corresponden a columnas incluidas en los rangos efectivos; no son registros externos a la configuración real.
- `TURNOS-PRECIOS`: `A1:F20` (120 posiciones, con 17 fórmulas excluidas de persistencia: 103 registros), `A22:F23` (12) y `B25:D31` (21) = **136 registros**.
- Total: **155 + 136 = 291**. La consulta no identificó coordenadas fuera de los seis rangos efectivos. No se leyeron campos de valores tarifarios ni se ejecutó ninguna escritura en Access.

Esta evidencia confirma el recuento ya declarado en los apartados 2.9 y 3.8 y sustituye la hipótesis anterior de 113 posiciones para `TARIFAS UNIDAD`.

#### 11.11. Rectificación del Lease Lock individual de producción — Orden 1L

La revisión del extracto literal de `SincronizarFilaHaciaAccess` (líneas 1737-1993 del módulo integrado) ha detectado afirmaciones residuales incorrectas en los apartados 3.4 y 3.7 y en la fila 8 del Anexo A, heredadas de la interpretación previa del Informe 1H.

Se rectifica que:

1. El `WHERE` de producción compara `UsuarioModificacion` con NULL, cadena vacía y el token de instancia `tokenMod`; **no** contiene la condición `UsuarioModificacion = uActual` ni un fallback de usuario `admin` en esa sentencia.
2. La firma real de `SincronizarFilaHaciaAccess` recibe `ws`, `fila` y `resTarget`; no tiene parámetro `permitirAlta`.
3. Cuando `recsLock <> 1`, la función consulta la clave, pero no distingue el resultado `EOF`: emite el mismo mensaje genérico de conflicto, revierte la transacción, llama a `SincronizarOrdenHaciaExcel(ws, numOrden, fila, resTarget)` y termina con `False`.
4. El bloque `rs.AddNew` se encuentra después de la adquisición satisfactoria (`recsLock = 1`) y de la apertura del Recordset; no es un mecanismo de recuperación del fallo de adquisición por clave inexistente.
5. `AsegurarOrdenExiste` es un helper separado; no se atribuye su invocación automática a `SincronizarFilaHaciaAccess` sin acreditar el llamador concreto.

Se actualizan las condiciones exactas para un bloqueo vigente (titular no nulo/vacío y distinto de `tokenMod`, marca temporal no nula y `DateDiff` entre 0 y 119), el tratamiento de fecha futura (`DateDiff < 0`) y la fila correspondiente del Anexo A. Se conserva la condición `DateDiff < 0` exclusiva de producción frente al SQL de laboratorio V5.2.

**Alcance:** rectificación documental limitada a estos predicados, escenarios y referencias de procedimiento. No altera `DEF-FIFO-07`, el catálogo de 18 defectos, ORD-10, TAR-10, la clasificación Certeza 4 ni la condición de candidata a revisión independiente. No se modifica código ni base de datos.

### ANEXO A — INVENTARIO EXHAUSTIVO DE OPERACIONES ADODB DE PRODUCCIÓN (32 PROCEDIMIENTOS DIRECTOS Y 2 ORQUESTADORES COMPUESTOS)
A continuación se detalla la matriz técnica completa y reconciliada del subsistema de persistencia en `modDatabase.bas`, integrado por **32 procedimientos ejecutores directos de sentencias DML/DDL** y **2 procedimientos orquestadores compuestos de persistencia** (`ArchivarOrdenesAntiguasBD` y `GestionarEntradaNumOrdenHoja`), sumando un total verificado de **34 procedimientos de persistencia** sobre el backend Access:

| Nº | Procedimiento | Líneas | Operación SQL / ADO | Tabla Afectada | Conexión | Transacción | Validaciones Previas | Manejo de Error | Llamadores Principales |
| :---: | :--- | :---: | :--- | :--- | :---: | :---: | :--- | :--- | :--- |
| 1 | `CambiarClave` | 438-452 | UPDATE | `Usuarios` | Dinámica ADO | No | Longitud hash SHA-256 | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 2 | `InsertarOrdenAtomica` | 846-976 | INSERT INTO | `Ordenes` | Dinámica ADO | Sí (`BeginTrans`) | Clave única, sonda red | `RollbackTrans` | Uso interno / Pruebas |
| 3 | `AsegurarEsquemaConcurrenciaBD` | 1118-1200 | ALTER TABLE, CREATE INDEX | `Ordenes`, `Ordenes_Historico` | Dinámica ADO | No | Comprobación de campos en esquema | `On Error Resume Next` | 14 llamadas internas |
| 4 | `IntentarReservarOrdenBD` | 1281-1401 | INSERT INTO | `Ordenes` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, lectura previa NumOrden | `RollbackTrans` | `GestionarEntradaNumOrdenHoja` |
| 5 | `ReservarSiguienteOrdenLibreBD` | 1403-1472 | INSERT INTO | `Ordenes` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, MAX < 900001 (bucle iter) | `RollbackTrans` | `GestionarEntradaNumOrdenHoja` |
| 6 | `RenumerarOrdenExistenteBD` | 1475-1525 | UPDATE | `Ordenes` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, existencia orden | `RollbackTrans` | `GestionarEntradaNumOrdenHoja` |
| 7 | `AsegurarOrdenExiste` | 1698-1732 | INSERT INTO | `Ordenes` | Dinámica ADO | No | SELECT COUNT(*) previo | Retorna Boolean | `ModuloSeguridadLocal.bas` |
| 8 | `SincronizarFilaHaciaAccess` | 1737-1993 | UPDATE condicional de adquisición + Recordset AddNew / Update | `Ordenes` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red; Lease Lock condicional con `tokenMod`; admite titular NULL/vacío, token propio, marca temporal NULL, antigüedad ≥120 s y marca futura (`DateDiff < 0`); **no** compara con `uActual` | Si `recsLock <> 1`, ejecuta SELECT, pero no distingue `EOF`: muestra el mensaje genérico de conflicto, revierte y refresca mediante `SincronizarOrdenHaciaExcel`; no tiene parámetro `permitirAlta`. El `AddNew` aparece después de una adquisición exitosa, no como recuperación de `recsLock = 0` | Hojas operativas: `Hoja2` (RESIDENCIA GIJÓN), `Hoja20` (RESIDENCIA SOTO) y `Hoja24` (RESIDENCIA OVIEDO), según correspondencia confirmada en el apartado 2.4 |
| 9 | `SincronizarLoteImportacionBD` | 2021-2238 | Recordset: AddNew / Update por fila | `Ordenes` | Dinámica ADO | Sí (`BeginTrans` para el lote) | Sonda red; selección por `Residencia`; **sin adquisición ni comprobación del Lease Lock por expediente** | `RollbackTrans` del lote ante error; `CommitTrans` al completar el bucle | `CopiarDatosHojasGijonSotoOviedo.bas` (llamador del inventario previo; no revalidado en Orden 1H) |
| 10 | `EjecutarUpdateCondicionalOrden` | 2448-2505 | UPDATE condicional | `Ordenes` | Dinámica ADO | No | Validación token sesión | Retorna Boolean | Procedimientos internos V5.2 |
| 11 | `ObtenerSiguienteNumFactura` | 2678-2778 | UPDATE | `ContadorFacturas` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, bloqueo exclusivo | `RollbackTrans` | `AsignarNumFactura.bas` |
| 12 | `SincronizarContadorFacturaManual`| 2782-2801 | UPDATE | `ContadorFacturas` | Dinámica ADO | No | Mayor que último asignado | `On Error Resume Next` | `AsignarNumFactura.bas` |
| 13 | `AnadirAListaNegra` | 3114-3155 | INSERT INTO / UPDATE | `ListaNegra` | Dinámica ADO | No | DNI o Nombre no vacío | Retorna Boolean | `frmGestionListaNegra.frm` |
| 14 | `RetirarDeListaNegra` | 3158-3166 | UPDATE (Activo=False) | `ListaNegra` | Dinámica ADO | No | DNI no vacío | Retorna Boolean | `frmGestionListaNegra.frm` |
| 15 | `InsertarLog` | 3176-3201 | INSERT INTO | `LogActividad` | Dinámica ADO | No | Usuario y operación presentes | Retorna Boolean | 11 módulos del sistema |
| 16 | `AltaUsuario` | 3930-3950 | INSERT INTO | `Usuarios` | Dinámica ADO | No | Existencia de usuario | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 17 | `BajaUsuario` | 3953-3963 | UPDATE (Activo=False)| `Usuarios` | Dinámica ADO | No | No ser usuario admin | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 18 | `AgregarResidenciaUsuario` | 3968-4045 | UPDATE | `Usuarios` | Dinámica ADO | No | Existencia y formato sedes | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 19 | `QuitarResidenciaUsuario` | 4050-4141 | UPDATE | `Usuarios` | Dinámica ADO | No | Validación sede previa | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 20 | `EliminarOrdenBD` | 4261-4420 | DELETE, INSERT INTO | `Ordenes`, `Ordenes_Historico` | Dinámica ADO | Sí (`BeginTrans`) | Lease Lock, datos reales | `RollbackTrans` | `BorradoNumOrden.bas` |
| 21 | `EliminarOrdenesLoteBD` | 4424-4552 | DELETE, INSERT INTO | `Ordenes`, `Ordenes_Historico` | Dinámica ADO | Sí (`BeginTrans`) | Colección > 0, Lease Lock | `RollbackTrans` | `BorradoSolicitudes.bas` |
| 22 | `EliminarTodasOrdenesSedeBD` | 4556-4659 | DELETE, INSERT INTO | `Ordenes`, `Ordenes_Historico` | Dinámica ADO | Sí (`BeginTrans`) | Clave sede válida, Lease Lock | `RollbackTrans` | `BorradoDatosAlojamiento.bas` |
| 23 | `PurgarHistoricoRGPD` | 4973-5003 | DELETE FROM | `Ordenes_Historico`, `Snapshot`| Dinámica ADO | No | Conteo previo registros purga | Retorna Long | `frmPanelRGPD.frm` |
| 24 | `AltaUsuarioDelegado` | 5209-5248 | INSERT INTO | `Usuarios` | Dinámica ADO | No | Validación formato y rol | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 25 | `AgregarResidenciaUsuarioDelegado`| 5251-5306| UPDATE | `Usuarios` | Dinámica ADO | No | Formato sede | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 26 | `QuitarResidenciaUsuarioDelegado` | 5309-5362 | UPDATE | `Usuarios` | Dinámica ADO | No | Formato sede | Retorna Boolean | `GestionarAccesoUsuarios.bas` |
| 27 | `AsegurarEsquemaBloqueosYHabitacionesBD`| 5716-5858| CREATE TABLE, ALTER TABLE| `BloqueosHabitaciones`, `EstadoHab`| Dinámica ADO| No | Comprobación existencia previa | `On Error Resume Next` | Procedimientos calendario |
| 28 | `GuardarBloqueosEnBD` | 5969-6110 | AddNew / Update / Delete | `BloqueosHabitaciones` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, lectura previa | `RollbackTrans` | `ThisWorkbook.cls` |
| 29 | `GuardarEstadoFilaCalendarioEnBD` | 6115-6183 | UPDATE / INSERT INTO | `EstadoHabitacionesCalendario` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red, clave sede | `RollbackTrans` | `ModuloGestionHabitaciones.bas`|
| 30 | `GuardarBloqueosYEstadoHabitacionesEnBD`| 6188-6306| AddNew / Update / Delete | `Bloqueos`, `EstadoHab` | Dinámica ADO | Sí (`BeginTrans`) | Sonda red | `RollbackTrans` | `CopiarDatosHojasGijonSotoOviedo.bas`|
| 31 | `AsegurarEsquemaTarifasBD` | 6746-6841 | CREATE TABLE, ALTER TABLE| `ConfigTarifasCeldas` | Dinámica ADO | No | Comprobación de campos en esquema | `On Error Resume Next` | Procedimientos tarifas |
| 32 | `GuardarTarifasEnBD` | 6894-7073 | AddNew / Update | `ConfigTarifasCeldas` | Dinámica ADO | Sí (`BeginTrans`) | Firma hash SHA-256 previa | `RollbackTrans` | `ThisWorkbook.cls` |

#### Procedimientos de Orquestación Compuesta de Persistencia (2 Procedimientos)
Existen dos procedimientos de alto nivel en `modDatabase.bas` que coordinan interfaz y persistencia sin ejecutar DML directo, delegando en los procedimientos anteriores:
1. **`ArchivarOrdenesAntiguasBD` (Líneas 4784-4970):**  
   - *Fase 1 (Persistencia Access):* Delega las mutaciones DML en `EliminarOrdenesLoteBD` (Entrada #21), que ejecuta la transacción atómica de transvase a `Ordenes_Historico` y borrado de `Ordenes`.
   - *Fase 2 (Compactación Excel):* Ejecuta la eliminación visual de filas mediante `Application.Union(...).Delete Shift:=xlUp` o `ClearContents`.
2. **`GestionarEntradaNumOrdenHoja` (Líneas 1527-1696):**  
   - Orquesta la respuesta interactiva ante entradas en celda A de hojas operativas, delegando el DML en `IntentarReservarOrdenBD` (Entrada #4), `RenumerarOrdenExistenteBD` (Entrada #6), `ReservarSiguienteOrdenLibreBD` (Entrada #5) y `SincronizarFilaHaciaAccess` (Entrada #8).
- **Conciliación Total:** 32 procedimientos ejecutores directos de sentencias DML/DDL + 2 procedimientos orquestadores de persistencia = **34 procedimientos involucrados en el ciclo de persistencia de `modDatabase.bas`**.

---
**FIN DEL DOCUMENTO MAESTRO DE ARQUITECTURA Y REQUISITOS — VERSIÓN 0.2.7**
