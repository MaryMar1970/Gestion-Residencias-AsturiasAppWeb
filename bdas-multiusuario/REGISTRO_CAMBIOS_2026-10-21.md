# REGISTRO DE CAMBIOS - 2026-10-21

## Hito 2026-10-21 00:55: Consolidación y Promoción a Versión de Referencia Oficial

### 1. Resumen Ejecutivo
Se ha promocionado formalmente la versión activa de trabajo `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando el hito previo (2026-10-20 23:15) en el histórico, actualizando el script de restauración inmediata (`Restaurar-Version-Referencia.ps1`), registrando exhaustivamente la trazabilidad de copias de seguridad incrementales y sincronizando la copia de despliegue en red.

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Atributos |
| :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 7.171.197 | `E8AE2E58DC520C55042E93C8992945F9` | Solo Lectura (`ReadOnly = True`) |
| **Versión de Referencia (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.725.632 | `E3586759300882BCEEEB1677F31643B7` | Solo Lectura (`ReadOnly = True`) |
| **Versión Activa de Trabajo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 7.171.197 | `E8AE2E58DC520C55042E93C8992945F9` | Lectura / Escritura |
| **Base de Datos Activa** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.725.632 | `E3586759300882BCEEEB1677F31643B7` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 7.171.197 | `E8AE2E58DC520C55042E93C8992945F9` | Lectura / Escritura |

#### Histórico Archivado del Hito Previo (2026-10-20 23:15):
- `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_2315.xlsm` (7.151.821 bytes, MD5: `5A23A4AB3A34A88442E6C28D5B1FB021`, `ReadOnly = True`)
- `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-20_2315.accdb` (6.725.632 bytes, MD5: `2D5E9D788825B17AAECEF46531D9ADFC`, `ReadOnly = True`)

---

### 3. Vinculación de Copias Incrementales con Funcionalidades y Código

Durante esta sesión técnica se generaron 5 copias de seguridad incrementales en `H:\ResidenciaBD\BackupsDiarios\`, vinculadas de forma biunívoca a cada funcionalidad desarrollada:

| Copia Incremental XLSM / ACCDB | Fecha / Hora | MD5 Criptográfico | Tamaño (bytes) | Funcionalidad y Cambios Asociados |
| :---| :---: | :---:| :---:| :---|
| `BDAS_Multiusuario_BACKUP_2026-10-20_2330.xlsm` | 20/10/2026 23:30 | `5A23A4AB3A34A88442E6C28D5B1FB021` | 7.151.821 | **Línea Base de Inicio:** Estado consolidado del Hito 2026-10-20 23:15 (Tarifas Dinámicas vinculadas a `TARIFAS UNIDAD`, Saneamiento de Hojas LOG y Sesión Persistente). |
| `BDAS_Multiusuario_BACKUP_2026-10-20_2341.xlsm` | 20/10/2026 23:41 | `1E3136455228CE5BC69C0F2C265D7B81` | 7.144.747 | **Confirmación de Salida Segura:** Modificación de `Workbook_BeforeClose` en `ThisWorkbook.cls` con mensaje interactivo (`¿Desea salir de la aplicación?`) para evitar pérdidas fortuitas ante el aspa superior (X). |
| `BDAS_Multiusuario_BACKUP_2026-10-20_2358.xlsm` | 20/10/2026 23:58 | `DB94E2E3AC6FCA5D52C348CD11404971` | 7.157.112 | **Cierre Suave UX y Gobernanza Anti-Mojibake:** Implementación de transición suave (Opción A: bloqueo de parpadeo con `ScreenUpdating = False`, minimización estética y guardado controlado). Saneamiento de textos ANSI en salida y adición del Protocolo Estricto Cero Mojibake a `AGENTS.md`. |
| `BDAS_Multiusuario_BACKUP_2026-10-21_0038.xlsm` | 21/10/2026 00:38 | `A04FBEC775090C7ECB5E26200EA9B42F` | 7.163.097 | **Sustitución de Flechas Corruptas:** Corrección en `EvitarDuplicidadSolicitudesGyS.bas` sustituyendo secuencias desalineadas `âž”` por `->` en las 3 opciones del `InputBox` de duplicados en Access. |
| `BDAS_Multiusuario_BACKUP_2026-10-21_0055_VERSION_REFERENCIA.xlsm` | 21/10/2026 00:55 | `E8AE2E58DC520C55042E93C8992945F9` | 7.171.197 | **Saneamiento Ortográfico Integral y Promoción a Referencia:** Reemplazo de todos los caracteres corruptos de acentuación (`ï¿½`, `peticin`, `vlida`, `OPCIN`, etc.) en `VerificarDuplicadoUltimos30Dias`. Compilación limpia 0 errores, escaneo 0 mojibake y promoción oficial. |
| `Residencia_BE_BACKUP_2026-10-21_0055_VERSION_REFERENCIA.accdb` | 21/10/2026 00:55 | `E3586759300882BCEEEB1677F31643B7` | 6.725.632 | **Copia Incremental de Base de Datos Backend:** Respaldo completo de la base de datos Access backend sincronizada con el hito de referencia. |

---

### 4. Alcance Técnico Consolidado en este Hito

1. **UX de Salida y Cierre Seguro de la Aplicación (`ThisWorkbook.cls`)**:
   - Diálogo interactivo ante el evento de cierre por aspa: `MsgBox("¿Desea salir de la aplicación?", vbQuestion + vbYesNo + vbDefaultButton2)`.
   - Transición fluida y elegante (Alternativa A nativa): se desactiva el parpadeo de pantalla con `Application.ScreenUpdating = False`, se minimiza la ventana (`xlMinimized`), se asegura el guardado ordenado si procede (`ThisWorkbook.Save`) y se finaliza la instancia de Excel si no existen otros libros abiertos.
2. **Gobernanza Técnica: Protocolo Estricto de Codificación y Cero Mojibake en VBA**:
   - Blindaje de 3 capas en `.agents/AGENTS.md`:
     - Capa 1: Codificación ANSI/Windows-1252 nativa obligatoria.
     - Capa 2: Prohibición expresa de scripts VBScript para textos en español.
     - Capa 3: Control de puerta pre-entrega obligatorio con escaneo automatizado.
3. **Saneamiento Ortográfico Integral en Módulo de Control de Duplicados (`EvitarDuplicidadSolicitudesGyS.bas`)**:
   - Saneamiento de todos los acentos corruptos y mojibake en el procedimiento `VerificarDuplicadoUltimos30Dias` (`últimos`, `número`, `DECLARACIÓN`, `comparación`, `numérico`, `petición`, `mayúsculas`, `Conversión`, `válido`, `válida`, `VALIDACIÓN`, `OPCIÓN`, `fórmulas`, `podrían`, etc.).
   - Sustitución de caracteres de flechas desalineados `âž”` por `->` en las 3 opciones del cuadro interactivo.
   - Preservación estricta del 100% de la lógica y flujos operativos existentes.
4. **Preservación Total de Componentes y Funcionalidades Previas**:
   - Tarifas dinámicas en facturas de Gijón y Oviedo vinculadas a `TARIFAS UNIDAD`.
   - Preservación total de los 171 componentes VBA.
   - Saneamiento integral de auditoría en hojas LOG (columnas normalizadas A–F).
   - Sincronización relacional inversa Top 500 y desambiguación de Login por sede.
   - Sesión persistente y auto-reparable en Nombres Definidos ocultos.
   - Erradicación de error 50290 en `Hoja11` (`SOBRE GIJÓN`).
   - Blindaje permanente de la barra de estado.
5. **Inmutabilidad Absoluta del Golden Master**:
   - `BDAS_Multiusuario_GOLDEN_MASTER.xlsm` permanece intacto, inalterado y protegido.

---

### 5. Protocolo de Reversión

En caso de requerirse en el futuro restaurar el entorno al estado óptimo de este hito:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
El script realiza un backup automático de seguridad previo y restablece tanto el libro Excel de trabajo como la base de datos backend Access y la copia de red.
