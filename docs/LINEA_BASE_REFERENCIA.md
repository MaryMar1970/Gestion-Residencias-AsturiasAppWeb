# Línea Base de Referencia Técnica — BDAS Multiusuario

> **Estado de Consolidación:** Activa y Verificada  
> **Fecha de Certificación:** 17 de Octubre de 2026 (00:15)  
> **Ámbito de Aplicación:** `H:\ResidenciaApp` (Código y Scripts) / `H:\ResidenciaBD` (Datos y Plantillas)

---

## 1. Certificado de Integridad y Huellas Digitales (Hashes)

Para asegurar que cualquier trabajo se realiza sobre una versión íntegra y trazable, se fijan las siguientes huellas criptográficas oficiales:

| Elemento | Ruta Absoluta | Hash MD5 Oficial | Estado |
| :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial y Trabajo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `F814E071502409B94A89C2FA545AC513` | 🟢 **REFERENCIA ACTIVA** (17/10/2026 23:16) |
| **Base de Datos Backend** | `H:\ResidenciaBD\Residencia_BE.accdb` | `740739C2D02DD8D98F17C4DD75F09AD2` | 🟢 Producción |
| **Libro Compartido en Red** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | `F814E071502409B94A89C2FA545AC513` | 📁 Servidor Red (Sincronizado) |
| **Último Backup Verificado** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-17_2316_POST_BLINDAJE_FORMULAS.xlsm` | `F814E071502409B94A89C2FA545AC513` | 🔒 Respaldo Blindaje Fórmulas |

---

## 2. Réplicas Maestras y Puntos de Restauración

Siguiendo las reglas de gobernanza del proyecto ([AGENTS.md](file:///H:/ResidenciaApp/.agents/AGENTS.md)), existen dos niveles de protección de referencia:

### A. Versión de Referencia Estable Activa (Hito 2026-10-11)
- **Ubicación protegida:**
  - `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  - `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
- **Características:** Incorpora el Selector Multirresidencia dinámico con Checkboxes, registro de auditoría en hojas LOG y RibbonX unificado.
- **Mecanismo de restauración inmediata:**
  Ejecutar el script [Restaurar-Version-Referencia.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Version-Referencia.ps1).

### B. Protocolo de Recuperación Instantánea (Golden Master)
- **Ubicación protegida (Solo Lectura estricto):**
  - `H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm`
- **Etiqueta en Git:** `GOLDEN_MASTER_TEXTOS_OK`
- **Inmutabilidad:** Esta réplica **NUNCA** se sobrescribe, quedando reservada exclusivamente para contingencias mayores de inicio de proyecto.
- **Mecanismo de restauración de emergencia:**
  Ejecutar el script [Restaurar-Golden-Master.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Golden-Master.ps1).

---

## 3. Inventario Físico de Componentes de Código (141 Archivos)

El directorio de módulos exportados [vba-modules](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules) contiene la totalidad del código fuente desagregado:

*   **Módulos Estándar (`.bas`):** 117 módulos.
    *   *Infraestructura Core:* `modDatabase.bas` (100 KB), `ModuloConfigSegura.bas` (25 KB), `GestionarAccesoUsuarios.bas`.
    *   *Lógica de Negocio:* `AsignarNumFactura.bas`, `FechasPeticion.bas`, `EvitarDuplicidadSolicitudesGyS.bas`, `DuplicadosOrdenyFactura.bas`.
    *   *Calendarios y Ocupación:* `ModuloCalendarioGijon.bas`, `ModuloCalendarioOviedo.bas`, `ModuloCalendarioSoto.bas`.
    *   *Facturación y Liquidación:* `FacturacionMesGIJON.bas`, `FacturacionMesOVIEDO.bas`, `FacturacionMesSOTO.bas`.
    *   *Lavandería:* `ModuloLavanderia_Core.bas`, `ModuloLavanderia_Gijon.bas`, `ModuloLavanderia_Oviedo.bas`, `ModuloLavanderia_Soto.bas`.
    *   *Ribbon y Botones:* 35 módulos de interconexión con la cinta de opciones.
*   **Módulos de Clase (`.cls`):**
    *   [Hoja34_TarifasUnidad.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja34_TarifasUnidad.cls): Controlador de eventos de hoja y validación anti-solapes en tiempo real.
*   **Formularios de Usuario (`.frm` y binarios `.frx`):** 11 formularios interactivos.
    *   `frmLogin`: Autenticación de usuario con hash SHA-256.
    *   `frmSelectorResidencia`: Selección dinámica de sedes asignadas.
    *   `frmPanelRGPD`: Panel de control de política de retención y borrado legal de datos.
    *   `frmGestionPasswords`: Panel de administración de contraseñas ofuscadas del sistema.
    *   `frmReservasUnidad`: Formulario principal de alta y edición de solicitudes.
    *   `frmBusqueda`: Búsqueda multicriterio.
    *   `frmConfigBusquedaBDAS`: Ajustes de búsqueda en BDAS.
    *   `frmCambiarPassword`: Cambio de clave personal por el usuario.
    *   `frmAviso`: Mensajes modales formateados.
    *   `frmSeparacion`: Configuración de margen de días entre reservas.
    *   `ResultadosGijon` / `ResultadosOviedo` / `ResultadosSoto`: Vistas de consulta rápida.

---

## 4. Esquema Consolidado de la Base de Datos Access (`Residencia_BE.accdb`)

La base de datos contiene 8 tablas estructuradas:

```
┌─────────────────────────────────┐       ┌─────────────────────────────────┐
│            Usuarios             │       │             Ordenes             │
├─────────────────────────────────┤       ├─────────────────────────────────┤
│ Id (PK)                         │       │ Id (PK)                         │
│ NombreUsuario                   │       │ NumOrden (Indexed)              │
│ NombreCompleto                  │       │ Residencia (GIJON, SOTO, OVIEDO)│
│ Clave (SHA-256)                 │       │ DNI, Nombre, Apellidos          │
│ Rol (ADMIN / OPERADOR)          │       │ FechaEntrada, FechaSalida       │
│ Residencias (CSV)               │       │ Importe, NumFactura, EstadoPago │
│ Activo (Boolean)                │       │ 41 campos en total...           │
└─────────────────────────────────┘       └─────────────────────────────────┘
                                                           ▲
                                                           │ (Archivado RGPD)
                                                           ▼
┌─────────────────────────────────┐       ┌─────────────────────────────────┐
│         ContadorFacturas        │       │        Ordenes_Historico        │
├─────────────────────────────────┤       ├─────────────────────────────────┤
│ Id (PK)                         │       │ Id (PK)                         │
│ Ejercicio                       │       │ NumOrden                        │
│ Residencia                      │       │ Residencia                      │
│ UltimoNumero                    │       │ 41 campos idénticos a Ordenes   │
└─────────────────────────────────┘       └─────────────────────────────────┘
```

*   **`Ordenes`:** 41 columnas. Almacena las órdenes activas en curso.
*   **`Ordenes_Historico`:** 41 columnas. Archivo histórico para auditoría y cumplimiento de retención RGPD.
*   **`Usuarios`:** 8 columnas. Credenciales y roles (`ADMIN` / `OPERADOR`).
*   **`ContadorFacturas`:** 4 columnas. Control atómico del correlativo de facturas para evitar duplicidades en entornos concurrentes.
*   **`ListaNegra`:** 6 columnas. Control de exclusiones.
*   **`LogActividad`:** 5 columnas. Trazabilidad de operaciones de administración.
*   **`CodigosPostales`:** 5 columnas. Catálogo nacional de municipios y códigos postales.
*   **`BDAS_Historico_Snapshot`:** 41 columnas. Copia de respaldo de seguridad de datos importados.
