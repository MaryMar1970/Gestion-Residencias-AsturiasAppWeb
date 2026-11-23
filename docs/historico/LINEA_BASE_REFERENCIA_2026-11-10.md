# Línea Base de Referencia Técnica — BDAS Multiusuario

> **Estado de Consolidación:** Activa y Verificada  
> **Fecha de Certificación:** 10 de Noviembre de 2026 (15:45 CET) — Hito 2026-11-10 15:45  
> **Ámbito de Aplicación:** `H:\ResidenciaApp` (Código y Scripts) / `H:\ResidenciaBD` (Datos y Plantillas)

---

## 1. Certificado de Integridad y Huellas Digitales (Hashes)

Para asegurar que cualquier trabajo se realiza sobre una versión íntegra y trazable, se fijan las siguientes huellas criptográficas oficiales:

| Elemento | Ruta Absoluta | Hash MD5 Oficial | Hash SHA-256 Oficial | Estado |
| :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial y Trabajo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | 🟢 **REFERENCIA ACTIVA** (Hito 2026-11-10 15:45) |
| **Base de Datos Backend** | `H:\ResidenciaBD\Residencia_BE.accdb` | `D9BB14ABB808DE0255E47F7A1AD3CC2C` | `EBCD61EBB64416A009F35F29C2A2B6F03E2CD10890CBAD2A55546949E214DCFD` | 🟢 Producción (Saneada y Conectada) |
| **Libro Compartido en Red** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | 🌐 Servidor Red (Sincronizado) |
| **Versión de Referencia Protegida** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Base de Datos Referencia Protegida** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | `D9BB14ABB808DE0255E47F7A1AD3CC2C` | `EBCD61EBB64416A009F35F29C2A2B6F03E2CD10890CBAD2A55546949E214DCFD` | 🔒 Solo Lectura (`IsReadOnly = True`) |
| **Último Backup Verificado** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-11-10_1545_VERSION_REFERENCIA.xlsm` | `DF455E2A5D9F215DD7E97DA8E0E335FC` | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | 🔒 Respaldo Consolidado Cierre Jornada |

---

## 2. Réplicas Maestras y Puntos de Restauración

Siguiendo las reglas de gobernanza del proyecto ([AGENTS.md](file:///H:/ResidenciaApp/.agents/AGENTS.md)), existen dos niveles de protección de referencia:

### A. Versión de Referencia Estable Activa (Hito 2026-11-10 15:45)
- **Ubicación protegida:**
  - `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  - `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
- **Características Clave del Hito:**
  1. **Blindaje de Celdas Combinadas:** Prevención de fallo por escritura en celdas combinadas `B25:D25` en `Turnos-Precios`.
  2. **Optimización de Carga:** Supresión de redibujado (`ScreenUpdating=False`) y recálculo automático (`Calculation=xlCalculationManual`) en `CargarTarifasDesdeBD`, acelerando la apertura a < 1 s.
  3. **Persistencia Multiusuario de Tarifas:** Esquema `ConfigTarifasCeldas` con 291 registros relacionales en Access e índice único libre de colisiones.
  4. **Cobertura Completa de Celdas:** Blindaje de rangos `E3:N9` para GIJÓN y `E13:N19` para OVIEDO, IVA `A3:C7` y turnos `A1:F20`, `A22:F23`, `B25:D31`.
  5. **Control Optimista de Concurrencia:** Firmas canónicas inyectivas deterministas por celda y manejo interactivo de colisiones en `Workbook_SheetDeactivate`.
  6. **Marco Consolidado Previo:** Bloqueos y Ocultación ACID V5.2, Resolución Dinámica `DB_PATH`, Saneamiento WAL SICO-N2, Protección `UserInterfaceOnly:=True`, Gestión RGPD y RibbonX.
- **Mecanismo de restauración inmediata:**
  Ejecutar el script [Restaurar-Version-Referencia.ps1](file:///H:/ResidenciaApp/bdas-multiusuario/scripts/Restaurar-Version-Referencia.ps1).

### B. Golden Master (Inmutable — Punto Cero 2026-10-24 00:55)
- **Ubicación:** `H:\ResidenciaBD\Plantilla\GoldenMaster\BDAS_Multiusuario_GOLDEN_MASTER.xlsm` y `Residencia_BE_GOLDEN_MASTER.accdb`.
- **Regla estricta:** Inmutable, reservado para reseteo completo del proyecto. Restaurable con `Restaurar-Golden-Master.ps1`.
