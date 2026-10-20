# REGISTRO DE CAMBIOS - 2026-10-20

## Hito 2026-10-20 23:15: Consolidación y Promoción a Versión de Referencia Oficial

### 1. Resumen Ejecutivo
Se ha promocionado formalmente la versión activa de trabajo a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando el hito previo de las 19:35 en el histórico, actualizando el script de restauración inmediata (`Restaurar-Version-Referencia.ps1`) y sincronizando la copia de despliegue en red.

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Atributos |
| :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 7.151.821 | `5A23A4AB3A34A88442E6C28D5B1FB021` | Solo Lectura (`ReadOnly`) |
| **Versión de Referencia (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.725.632 | `2D5E9D788825B17AAECEF46531D9ADFC` | Solo Lectura (`ReadOnly`) |
| **Versión Activa de Trabajo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 7.151.821 | `5A23A4AB3A34A88442E6C28D5B1FB021` | Lectura / Escritura |
| **Base de Datos Activa** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.725.632 | `2D5E9D788825B17AAECEF46531D9ADFC` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 7.151.821 | `5A23A4AB3A34A88442E6C28D5B1FB021` | Lectura / Escritura |

#### Histórico Archivado del Hito Previo (19:35):
- `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_1935.xlsm` (7.152.489 bytes, `ReadOnly`)
- `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-20_1935.accdb` (6.721.536 bytes, `ReadOnly`)

---

### 3. Alcance Técnico Consolidado en este Hito

1. **Tarifas Dinámicas en Facturas PAX (Gijón y Oviedo)**:
   - Enlace dinámico directo de importes con la hoja `TARIFAS UNIDAD`.
   - Garantía de cuadre contable exacto al céntimo en facturas PAX.
2. **Dictamen Forense de Arquitectura: Preservación Total e Incondicional de los 171 Componentes VBA**:
   - Auditoría de dependencias en las 27 hojas del libro (`Hoja2` [Gijón], `Hoja20` [Soto], `Hoja24` [Oviedo], etc.).
   - Constatación de acoplamientos internos críticos: los eventos `Worksheet_Change` de las hojas operativas referencian directamente variables públicas de `ModuloVariablesGlobales.bas` (`UltimoOrdenLiberado`, `UltimasHabitacionesLiberadas`, `UltimaHojaLiberada`), llamadas a `ModListaNegra.CheckBlacklist`, invocaciones a `frmConfigBusquedaBDAS.Show` y cálculo en `CalculoDiasUso.CalcularDiasUso`.
   - Se ratifica la directiva técnica: **No eliminar ningún módulo del proyecto**. Los 171 componentes deben permanecer intactos para mantener el 100% de operatividad y 0 errores de compilación.
3. **Mecanismo de Asignación y Auditoría de Facturas**:
   - Contador atómico en backend Access (`ContadorFacturas`, `modDatabase.ObtenerSiguienteNumFactura`) para garantizar estricta concurrencia monótona creciente (`UltimoNumero + 1`) sin solapamiento entre puestos.
   - Auditor de huecos en hoja activa (`ValidarFacturaManual`) que advierte sobre discontinuidades respecto a la serie registrada.
4. **Saneamiento Integral de Auditoría en Hojas LOG**:
   - Parseo estricto por columnas normalizadas: `[Usuario, Fecha/Hora, Nº Orden, Columna, Valor anterior, Valor nuevo]`.
   - Protección contra inyección de fórmulas (`'=` delante de expresiones de auditoría).
5. **Sincronización Relacional Top 500 y Blindaje Multirresidencia**:
   - Subconsulta descendente en Access para traer los 500 registros más recientes sin pérdida histórica.
   - Filtrado de Login exclusivo por sede para neutralizar contaminación cruzada.
6. **Resiliencia de Sesión y Blindaje de UI**:
   - Sesión persistente y auto-reparable en nombres definidos ocultos de Excel, inmune a reseteos de variables globales en VBA.
   - Erradicación de error 50290 en `Worksheet_Calculate` de `Hoja11` (`SOBRE GIJÓN`).
   - Blindaje permanente de la barra de estado con la leyenda de conexión.
   - Saneamiento global de mojibake y normalización ASCII 7-bit en búsquedas.
7. **Inmutabilidad Absoluta del Golden Master**:
   - `BDAS_Multiusuario_GOLDEN_MASTER.xlsm` permanece intacto, inalterado y protegido.

---

### 4. Protocolo de Reversión
En caso de requerirse en el futuro restaurar el entorno al estado óptimo de este hito:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
El script realiza un backup automático de seguridad previo y restablece tanto el libro Excel de trabajo como la base de datos backend Access y la copia de red.
