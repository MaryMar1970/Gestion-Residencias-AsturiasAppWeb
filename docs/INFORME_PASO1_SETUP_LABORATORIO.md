# RESIDENCIAAPP — PUERTA 2: LABORATORIO TÉCNICO
## INFORME OFICIAL DEL PASO 1 — CREACIÓN DEL ENTORNO, CLONACIÓN VERIFICADA Y AISLAMIENTO DE RUTAS

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Puerta 2 — Laboratorio Técnico v1.0  
**Hito autorizado:** PASO 1 — Creación de entorno, clonación y aislamiento de rutas  
**Naturaleza:** Preparación de infraestructura, control criptográfico y diagnóstico no destructivo  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha y hora:** 20 de noviembre de 2026 | 16:02  

---

## 1. COMPROBACIÓN PREVIA DE ESTADO QUIESCENTE (CONDICIÓN 1.2)

Inmediatamente antes de proceder a la lectura y copia de los archivos de origen, se ejecutó una verificación no invasiva del entorno operativo:

| Elemento Comprobado | Procedimiento Técnico | Resultado Obtenido | Estado |
| :--- | :--- | :--- | :---: |
| **Instancias activas de Excel** | Inspección de tabla de procesos (`tasklist`) | **0 procesos** `excel.exe` en memoria. | **PASS** |
| **Instancias activas de Access** | Inspección de tabla de procesos (`tasklist`) | **0 procesos** `msaccess.exe` en memoria. | **PASS** |
| **Fichero de bloqueo Access (`.laccdb`)** | Comprobación de existencia en `H:\ResidenciaBD\` | Inexistente (`Residencia_BE.laccdb` = `False`). | **PASS** |
| **Bloqueos residuales en directorio BD** | Escaneo de extensiones `*.laccdb` en `H:\ResidenciaBD\` | 0 ficheros de bloqueo detectados. | **PASS** |

**Dictamen de seguridad:** Estado 100% quiescente, consistente y sin conexiones abiertas de usuario ni de motor OLEDB/DAO.

---

## 2. RUTAS Y ESTRUCTURA DEL LABORATORIO AISLADO

En estricto cumplimiento del apartado 1.4 de la orden, el laboratorio se ha creado en un directorio exclusivo, completamente **fuera del repositorio Git** y sin activar la ruta reservada `C:\RESIDENCIAAPP`:

- **Directorio Raíz del Laboratorio:**  
  [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/)
- **Rama BASELINE (Inmutable, Solo Lectura):**  
  [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\BASELINE\`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/BASELINE/)
- **Rama TRABAJO (Entorno de Ensayos y Escritura):**  
  [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/TRABAJO/)
- **Rama EVIDENCIAS (Manifiestos y Logs):**  
  [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\EVIDENCIAS\`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/EVIDENCIAS/)

---

## 3. TRAZABILIDAD CRIPTOGRÁFICA Y CONTROL DE INTEGRIDAD

Se calcularon los hashes SHA-256 de los ficheros de origen antes de la copia, inmediatamente después de la copia, y se contrastaron bit a bit contra las réplicas en `BASELINE` y `TRABAJO`:

### 3.1. Base de Datos Access (`Residencia_BE.accdb`)

| Instancia | Ubicación Física Absoluta | Tamaño (bytes) | Hash SHA-256 | Estado |
| :--- | :--- | :---: | :--- | :---: |
| **ORIGEN (Pre-copia)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7.639.040 | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | Línea Base |
| **ORIGEN (Post-copia)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 7.639.040 | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **INTACTO** |
| **BASELINE (+R)** | `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\BASELINE\Residencia_BE.accdb` | 7.639.040 | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **IDÉNTICO** |
| **TRABAJO** | `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\Residencia_BE.accdb` | 7.639.040 | `77DB1D7FFEBA4184451C0B9496DFA9D7EEC27FFA3AE4C4A733686A042CBFDC03` | **IDÉNTICO** |

### 3.2. Libro Excel (`BDAS_Multiusuario.xlsm`)

| Instancia | Ubicación Física Absoluta | Tamaño (bytes) | Hash SHA-256 | Estado |
| :--- | :--- | :---: | :--- | :---: |
| **ORIGEN (Pre-copia)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | Línea Base |
| **ORIGEN (Post-copia)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **INTACTO** |
| **BASELINE (+R)** | `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\BASELINE\BDAS_Multiusuario.xlsm` | 6.621.963 | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **IDÉNTICO** |
| **TRABAJO** | `H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\BDAS_Multiusuario.xlsm` | 6.621.963 | `B9117659842369515FFAF714AB63B8BB3B1927CD2F7FB6D589882AC7F6E419D7` | **IDÉNTICO** |

**Conclusiones de Integridad:**
1. Los archivos operativos de origen en `H:\ResidenciaBD\` y en el repositorio Git no sufrieron ninguna alteración durante el copiado (hashes antes/después 100% coincidentes).
2. Las réplicas maestras en `BASELINE` son idénticas bit a bit a los originales y se les aplicó el atributo de Solo Lectura del sistema de archivos (`attrib +r`).
3. El manifiesto JSON de control quedó registrado en [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\EVIDENCIAS\manifest_paso1.json`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/EVIDENCIAS/manifest_paso1.json).

---

## 4. DEMOSTRACIÓN ESTÁTICA Y EFECTIVA DE AISLAMIENTO DE RUTA (`DB_PATH`)

### 4.1. Demostración Estática

En [`modDatabase.bas`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L63-L111), la propiedad pública `DB_PATH` implementa la siguiente lógica de cascada:

```vba
Public Property Get DB_PATH() As String
    If Len(m_sRutaCacheBD) > 0 Then
        DB_PATH = m_sRutaCacheBD
        Exit Property
    End If
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Prioridad 1: Carpeta local del libro activo (si no es C:\ResidenciaApp)
    Dim sRutaLocal As String
    If Len(ThisWorkbook.Path) > 0 Then
        sRutaLocal = ThisWorkbook.Path & "\Residencia_BE.accdb"
        If UCase$(ThisWorkbook.Path) <> "C:\RESIDENCIAAPP" And fso.FileExists(sRutaLocal) Then
            m_sRutaCacheBD = sRutaLocal
            DB_PATH = m_sRutaCacheBD
            Set fso = Nothing
            Exit Property
        End If
    End If
    ...
```

**Evaluación determinista para el libro en `TRABAJO`:**
1. `ThisWorkbook.Path` = `"H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO"` ($\text{longitud} = 58 > 0$).
2. `sRutaLocal` = `"H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\Residencia_BE.accdb"`.
3. `UCase$(ThisWorkbook.Path) <> "C:\RESIDENCIAAPP"` evalúa a **`True`**.
4. `fso.FileExists(sRutaLocal)` evalúa a **`True`** (la base clonada existe en la misma carpeta).
5. Se asigna `m_sRutaCacheBD = sRutaLocal`, se devuelve el valor y se ejecuta **`Exit Property`**.
6. **Conclusión estática:** Las ramas de Prioridad 2 (Preproducción de red) y Prioridad 3 (Desarrollo local) quedan completamente inaccesibles e inalcanzables.

---

### 4.2. Comprobación Efectiva en Tiempo de Ejecución (Runtime)

Para validar empíricamente la deducción estática sin desencadenar efectos colaterales de interfaz ni de base de datos, se automatizó Excel 2016 de 32 bits bajo el siguiente protocolo no invasivo:
- `Application.AutomationSecurity = 1` (`msoAutomationSecurityLow`, habilita macros).
- `Application.EnableEvents = False` (impide la ejecución de `Workbook_Open`, evitando login modal y logs).
- `Application.DisplayAlerts = False`.
- Apertura de [`H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\BDAS_Multiusuario.xlsm`](file:///H:/ResidenciaApp/LABORATORIO_GATE2_20261120_1558/TRABAJO/BDAS_Multiusuario.xlsm) en modo `ReadOnly = True`.
- Invocación de la función pública [`modDatabase.GetConnectionString`](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L682-L684) mediante `xl.Run("GetConnectionString")`.

**Salida real observada del motor VBA:**
```text
Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\ResidenciaApp\LABORATORIO_GATE2_20261120_1558\TRABAJO\Residencia_BE.accdb;
```

**Dictamen de Aislamiento:**  
Queda empírica y formalmente demostrado que el libro de `TRABAJO` conecta exclusivamente con la base de datos clonada de su propio directorio de laboratorio, sin interactuar en ningún momento con `H:\ResidenciaBD\Residencia_BE.accdb` ni con la red.

---

## 5. INCIDENCIAS, LIMITACIONES Y ESTADO DE CIERRE

- **Incidencias técnicas:** **0 incidencias**.
- **Ficheros `.laccdb` residuales:** **0** (cerrado limpio de Excel tras el diagnóstico).
- **Límites de alcance vigentes:**
  - `BASELINE` permanece cerrada, inmutable y protegida como Solo Lectura (`+R`).
  - No se ha ejecutado ninguna prueba funcional, transaccional, DDL ni DML de las previstas para los pasos posteriores de la Puerta 2.
  - Las estaciones remotas continúan clasificadas como **NO INSPECCIONADAS**.

---

### Decisión y Detención

El **PASO 1** ha concluido con éxito y con todas las garantías de seguridad técnica acreditadas.  
**El inspector técnico detiene su actuación en este punto**, a la espera de la revisión de la Dirección Arquitectónica y de la autorización expresa para iniciar las pruebas técnicas del **PASO 2**.