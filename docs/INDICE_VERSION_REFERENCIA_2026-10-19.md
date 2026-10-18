# Índice Técnico de la Versión de Referencia Oficial — Hito 2026-10-19 01:50

> **Ubicación de Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`  
> **Libro Maestro:** `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (MD5: `54d94f60f542de5210795c3c30a5a419`)  
> **Base de Datos Maestro:** `Residencia_BE_VERSION_REFERENCIA.accdb` (MD5: `ede9c7ecb540604f835facc39f8f08b8`)  
> **Estado de Atributos:** Ambos archivos protegidos contra escritura (`IsReadOnly = True`).  
> **Script de Recuperación:** `bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1`  

---

## 1. Alcance Funcional Consolidado en este Hito

Esta versión de referencia consolida la totalidad de mejoras del hito previo (2026-10-18 21:25) más las siguientes funcionalidades y correcciones críticas implementadas en la sesión:

1. **Carátula de Presentación Institucional a Pantalla Completa (Hoja `INICIO`)**:
   - Hoja `INICIO` configurada sin cuadrículas (`DisplayGridlines = False`), sin encabezados (`DisplayHeadings = False`), área de scroll bloqueada (`A1:BA45`) y fondo blanco unificado.
   - Fotografía histórica incrustada y embebida de forma nativa (`SaveWithDocument:=True`) de la Casa Cuartel ([Foto Casa Cuartel antigua (3).png](file:///H:/ResidenciaApp/bdas-multiusuario/Foto%20Casa%20Cuartel%20antigua%20(3).png), 1177.5 x 633.75 pt).
   - Renderizado inmediato en la apertura de Excel antes de ejecutar código VBA.
   - Formulario de autenticación modal `frmLogin` perfectamente centrado entre los dos guardias civiles de la imagen.
   - Transición fluida a la hoja de la residencia autorizada con auto-ocultación de `INICIO` (`wsInicio.Visible = xlSheetHidden`).

2. **Persistencia Universal e Inmediata de Reservas (ADODB UPSERT)**:
   - Rutinas centrales `modDatabase.AsegurarOrdenExiste` y `modDatabase.SincronizarFilaHaciaAccess`.
   - Guardado automático a nivel de fila disparado por `Worksheet_Change`, `Worksheet_Deactivate` y formularios de asignación de habitación (`SeleccionHabitacionGijon`, `SeleccionHabitacionSOTO`, `SeleccionHabitacionOVIEDO`).
   - Trasvase garantizado de fechas hacia las hojas calendario y conservación absoluta de datos al navegar o cambiar de vista.

3. **Sincronización Centralizada de Auditoría LOG**:
   - Rutina `ModuloLOG.SincronizarHojaLOGDesdeAccess` que descarga en tiempo real los registros auditados desde la tabla `LogActividad` de Access directamente a las hojas `LOG_GIJÓN`, `LOG_SOTO` y `LOG_OVIEDO`.

4. **Blindaje de Visibilidad de las Hojas de Gijón**:
   - Normalización de codificación de caracteres en [ThisWorkbook.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ThisWorkbook.cls) y [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas) (`GIJÓN` / `GIJON`).
   - Evita la ocultación involuntaria de las hojas operativas de Gijón tras abandonar los calendarios.

5. **Asignación Automática de "SIN TURNO" por Defecto en Soto**:
   - En la hoja `RESIDENCIA SOTO`, al obtener o ingresar una evaluación en la columna G (`EVALUACIÓN`), se asigna automáticamente por defecto `"SIN TURNO"` en la columna H (`TURNO`) si estaba vacía.
   - Conservación completa de la validación de lista desplegable para permitir al usuario cambiar a cualquier turno oficial (`S. Santa 1..3`, `Navidad 1..3`, `Verano 1..11`) en cualquier momento.
   - Sincronización automática del turno seleccionado hacia Access `Ordenes.Turno`.

6. **Eliminación Definitiva de la Hoja `LISTA NEGRA`**:
   - Gestión 100% en memoria en Access (`Residencia_BE.accdb`, tabla `ListaNegra`).
   - Desacople total de las hojas de Excel y consulta mediante diccionario de alto rendimiento en [frmGestionListaNegra.frm](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/frmGestionListaNegra.frm).

---

## 2. Inventario de Componentes y Estado

- **Módulos Estándar (.bas):** 52 módulos (`modDatabase`, `ModuloLOG`, `ModuloConfigSegura`, `BotonListaNegra`, etc.).
- **Formularios (.frm):** 9 formularios (`frmLogin`, `frmSelectorResidencia`, `frmGestionListaNegra`, `SeleccionHabitacionGijon`, `SeleccionHabitacionSOTO`, `SeleccionHabitacionOVIEDO`, etc.).
- **Módulos de Documento (.cls):** `ThisWorkbook`, `Hoja2` (Gijón), `Hoja20` (Soto), `Hoja24` (Oviedo), `Hoja8` (Estudiantes), `Hoja34` (Tarifas).
- **Hojas Excel:** 47 hojas (hojas operativas visibles según rol y residencia autenticada).

---

## 3. Protocolo de Reversión

En caso de cualquier eventualidad o necesidad de recuperar este estado estable:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
El script generará una copia de seguridad del estado de trabajo y restaurará de forma íntegra e instantánea los archivos maestro protegidos.
