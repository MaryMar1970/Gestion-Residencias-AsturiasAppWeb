# RESIDENCIAAPP — INFORME DE EVIDENCIAS TÉCNICAS (ORDEN 1X)
## VERIFICACIÓN DE RANGOS DE DESPLAZAMIENTO, INSERCIÓN, ORDENACIÓN Y LIMPIEZA DE FILAS OPERATIVAS

**Proyecto:** ResidenciaApp – BDAS Multiusuario  
**Fase:** Cierre del diseño arquitectónico de numeración, facturación y archivado  
**Naturaleza:** Inspección estática y extracción de código fuente, estrictamente de solo lectura  
**Dirección arquitectónica:** ChatGPT  
**Inspector técnico ejecutor:** Antigravity  
**Fecha de emisión:** 19 de noviembre de 2026  
**Documento de referencia:** Órdenes 1S, 1T, 1V y 1W  

---

## 1. CONTROL DE FUENTES, COMMIT Y HASHES

La presente verificación técnica se ha llevado a cabo de forma estrictamente pasiva y en modo solo lectura sobre las fuentes oficiales vigentes del repositorio y el libro activo.

### 1.1. Identificación del Commit de Trabajo
- **Rama:** `master`
- **Commit HEAD:** `9f67b2c4e1dad97e8c929163d936302b0815b665`

### 1.2. Integridad de los Archivos Fuente Examinados

| Archivo | Ruta Absoluta | Tamaño (Bytes) | SHA-256 |
| :--- | :--- | :---: | :--- |
| **Libro Activo** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.621.963 | `b9117659842369515ffaf714ab63b8bb3b1927cd2f7fb6d589882ac7f6e419d7` |
| **Borrado Solicitudes** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoSolicitudes.bas` | 21.439 | `c21411f02719bc0a7730e3fbb772ba8a6c2c59b8a5e512fd1dc1d5e173a646c1` |
| **Módulo Database** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\modDatabase.bas` | 298.794 | `ef4fdb01ae65e46129057903bef0df8125a398cac39a305b4e92438043c3ba47` |
| **Borrado Alojamiento** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BorradoDatosAlojamiento.bas` | 11.065 | `8910cfe366c08a3464caa2328460f24926a9ec8bc6c9051c20f2fc9f34fd20a0` |
| **Filtro Fechas** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\BotonFiltrarSolicitudesxFechas.bas` | 8.878 | `ef7ea90c4c47b5ae15a31eb2786937cb113fcab49f3238612f0f46dff21eb835` |
| **Módulo Seguridad Local** | `H:\ResidenciaApp\bdas-multiusuario\vba-modules\ModuloSeguridadLocal.bas` | 51.523 | `7486f0ff528c11e3b5e40e6983792cb0d09995c739bf824bf28cf0810aa157ff` |

---

## 2. COMPROBACIÓN A — BORRADO CON DESPLAZAMIENTO VERTICAL `Shift:=xlUp`

Se han inspeccionado los dos procedimientos indicados en el requerimiento técnico para acreditar la extensión horizontal exacta del desplazamiento y su afección a las columnas funcionales (`AD` en GIJÓN; `AD` en SOTO; `AC, AD, AE` en OVIEDO) y a la columna candidata de identidad `AJ` (columna 36).

### 2.1. Procedimiento 1: `BorrarSolicitudesPorOrden`
- **Archivo:** [BorradoSolicitudes.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoSolicitudes.bas#L175-L465)
- **Líneas del bloque de borrado:** 440 a 454.
- **Hojas sobre las que opera:** `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO` y `RESIDENCIA OVIEDO`.

#### 1. Código Fuente del Bloque de Localización y Eliminación
```vba
    ' Agrupar todas las filas seleccionadas en un único objeto Range mediante Union
    For Each fila In filasAEliminar
        If rngEliminar Is Nothing Then
            Set rngEliminar = ws.Rows(CLng(fila))
        Else
            Set rngEliminar = Application.Union(rngEliminar, ws.Rows(CLng(fila)))
        End If
    Next fila
    
    ' Eliminar físicamente en una sola llamada nativa (las filas inferiores suben de inmediato)
    If Not rngEliminar Is Nothing Then
        rngEliminar.Delete Shift:=xlUp
        Set rngEliminar = Nothing
    End If
```

#### 2. Instrucción Exacta de Borrado
```vba
rngEliminar.Delete Shift:=xlUp
```

#### 3. Referencias y Delimitación del Rango
- **Referencia utilizada:** `ws.Rows(CLng(fila))`.
- **Tipo de rango:** **Fila completa nativa de Excel (`Range.Rows`)**.
- **Extensión horizontal exacta:** Columnas 1 a 16.384 (rango completo `A:XFD`). No se delimita un subconjunto de columnas.

#### 4. Evaluación de Afección y Desplazamiento

| Hoja | Columna Funcional 1 | Columna Funcional 2 | Columna Funcional 3 | Identidad Candidata (AJ = 36) | Comportamiento Acreditado |
| :--- | :---: | :---: | :---: | :---: | :--- |
| **GIJÓN** | `AD` (30) | — | — | `AJ` (36) | **Se elimina y desplaza junto con el resto de la fila**. Al borrarse la fila completa (`ws.Rows`), las celdas de las filas inferiores en las columnas `AD` y `AJ` ascienden de forma perfectamente alineada y sincronizada con el resto de los datos del expediente. |
| **SOTO** | `AD` (30) | — | — | `AJ` (36) | **Se elimina y desplaza junto con el resto de la fila**. Las celdas inferiores ascienden conjuntamente en todas las columnas. |
| **OVIEDO** | `AC` (29) | `AD` (30) | `AE` (31) | `AJ` (36) | **Se eliminan y desplazan junto con el resto de la fila**. Toda la fila física se destruye y el bloque inferior sube de forma homogénea en las 16.384 columnas. |

---

### 2.2. Procedimiento 2: `ArchivarOrdenesAntiguasBD`
- **Archivo:** [modDatabase.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4784-L4950)
- **Líneas del bloque de borrado:** 4866 a 4901.
- **Hojas sobre las que opera:** `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO` y `RESIDENCIA OVIEDO`.

El procedimiento presenta **dos rutas de ejecución mutuamente excluyentes** según el volumen de filas a archivar:

#### 1. Código Fuente Completo del Bloque Dual

```vba
        Dim ultFilaA As Long
        ultFilaA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        Dim totalFilasHoja As Long
        totalFilasHoja = IIf(ultFilaA >= 2, ultFilaA - 1, 0)

        If colFilas.Count >= totalFilasHoja And ultFilaA >= 2 Then
            ' Si se archivan todas las filas de la hoja, limpiar contenidos preservando fórmulas de plantilla
            Select Case claveRes
                Case "GIJON"
                    ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 20)).ClearContents
                    ws.Range(ws.Cells(2, 22), ws.Cells(ultFilaA, 27)).ClearContents
                    ws.Range(ws.Cells(2, 30), ws.Cells(ultFilaA, 30)).ClearContents
                Case "SOTO"
                    ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 16)).ClearContents
                    ws.Range(ws.Cells(2, 19), ws.Cells(ultFilaA, 20)).ClearContents
                    ws.Range(ws.Cells(2, 22), ws.Cells(ultFilaA, 27)).ClearContents
                    ws.Range(ws.Cells(2, 30), ws.Cells(ultFilaA, 30)).ClearContents
                Case "OVIEDO"
                    ws.Range(ws.Cells(2, 1), ws.Cells(ultFilaA, 21)).ClearContents
                    ws.Range(ws.Cells(2, 23), ws.Cells(ultFilaA, 31)).ClearContents
            End Select
            ws.Range("P2:P" & ultFilaA).ClearComments
        Else
            ' Agrupar en bloques de hasta 100 filas para evitar el límite de áreas en Application.Union
            Dim rngBatch As Range, idxF As Long, contAreas As Long
            For idxF = colFilas.Count To 1 Step -1
                If rngBatch Is Nothing Then
                    Set rngBatch = ws.Rows(CLng(colFilas(idxF)))
                    contAreas = 1
                Else
                    Set rngBatch = Application.Union(rngBatch, ws.Rows(CLng(colFilas(idxF))))
                    contAreas = contAreas + 1
                End If
                If contAreas >= 100 Or idxF = 1 Then
                    rngBatch.Delete Shift:=xlUp
                    Set rngBatch = Nothing
                    contAreas = 0
                End If
            Next idxF
        End If
```

#### 2. Análisis por Separado de las Dos Rutas

##### Ruta A: Vaciado Total (`colFilas.Count >= totalFilasHoja`)
- **Instrucción:** No utiliza `Delete Shift:=xlUp`; ejecuta `ClearContents` por rangos de columnas específicos.
- **Extensión horizontal:**
  - GIJÓN: Cols 1..20, 22..27 y Col 30 (`AD`).
  - SOTO: Cols 1..16, 19..20, 22..27 y Col 30 (`AD`).
  - OVIEDO: Cols 1..21 y 23..31 (`W:AE`, incluye `AC, AD, AE`).
- **Afección a columnas funcionales:** Las columnas funcionales **se limpian** de forma efectiva.
- **Afección a columna candidata AJ (36):** **QUEDA FUERA DEL RANGO DE LIMPIEZA**. La columna `AJ` **permanece intacta**, conservando indebidamente sus datos locales pese a haberse archivado todos los expedientes de la hoja.

##### Ruta B: Borrado Parcial con Desplazamiento Vertical (`Else`)
- **Instrucción exacta:** `rngBatch.Delete Shift:=xlUp` en bloques de hasta 100 áreas ([modDatabase.bas#L4895](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L4895)).
- **Referencia utilizada:** `ws.Rows(CLng(colFilas(idxF)))`.
- **Tipo de rango:** **Fila completa nativa de Excel (`Range.Rows`)**.
- **Extensión horizontal:** Columnas 1 a 16.384 (`A:XFD`).
- **Afección a columnas funcionales:**
  - GIJÓN (`AD`): Se elimina y desplaza junto con toda la fila.
  - SOTO (`AD`): Se elimina y desplaza junto con toda la fila.
  - OVIEDO (`AC, AD, AE`): Se eliminan y desplazan junto con toda la fila.
- **Afección a columna candidata AJ (36):** **Se elimina y desplaza junto con toda la fila**, ascendiendo las filas inferiores en perfecta sincronía.

---

### 2.3. Procedimiento Complementario: `PurgarOperacionesBorradasDeHojas`
- **Archivo:** [ModuloSeguridadLocal.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L850-L880)
- **Líneas:** 850 a 882.
- **Contexto:** Purga de filas en las hojas operativas tras procesar borrados offline en la reconciliación SICO-N2.
- **Instrucción:** `rngDelSede.Delete Shift:=xlUp` (L880), construido sobre `wsBorr.Rows(fSede + 1)` (L851-L853).
- **Comportamiento:** Idéntico al caso de `ArchivarOrdenesAntiguasBD`: borrado de fila completa (1 a 16.384) en la ruta parcial, desplazando conjuntamente columnas funcionales y columna candidata `AJ`.

---

## 3. COMPROBACIÓN B — ORDENACIONES, FILTROS E INSERCIONES

### 3.1. Localización de Operaciones de Ordenación en Código Fuente

Se realizó una búsqueda exhaustiva en todo el código fuente VBA sobre patrones de ordenación:
- `.Sort`
- `SortFields`
- `Sort.SetRange`
- `AutoFilter.Sort`
- `Range.Sort`

#### Resultado de la Búsqueda
- **Coincidencias totales en el proyecto:** 1 única coincidencia ([ModuloSeguridadLocal.bas#L946](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloSeguridadLocal.bas#L946)), la cual ordena exclusivamente la hoja técnica de la cola local `SYS_COLA_SYNC` (`wsCola.Range(...).Sort`).
- **Coincidencias en hojas operativas (`RESIDENCIA GIJÓN`, `RESIDENCIA SOTO`, `RESIDENCIA OVIEDO`):** **CERO (0)**.
- **Acreditación técnica:** **No existe ninguna instrucción VBA programada que ejecute ordenaciones sobre las tres hojas operativas de residencia.** Las filas no se ordenan programáticamente por código en ningún punto del ciclo de vida de la aplicación.

---

### 3.2. Rangos de AutoFilter en el Libro y Extensión OpenXML

Se inspeccionó directamente la estructura interna OpenXML del libro activo `BDAS_Multiusuario.xlsm`:
- `xl/workbook.xml` (definiciones de nombres `_FilterDatabase` y hojas).
- `xl/worksheets/sheet2.xml` (`RESIDENCIA GIJÓN`).
- `xl/worksheets/sheet23.xml` (`RESIDENCIA SOTO`).
- `xl/worksheets/sheet34.xml` (`RESIDENCIA OVIEDO`).
- Módulo Ribbon: [BotonFiltrarSolicitudesxFechas.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonFiltrarSolicitudesxFechas.bas#L148).

#### 1. Definiciones OpenXML Extraídas

##### A. Rango `_xlnm._FilterDatabase` en `xl/workbook.xml`
```xml
<definedName name="_xlnm._FilterDatabase" localSheetId="1" hidden="1">'RESIDENCIA GIJÓN'!$A$1:$AB$1</definedName>
<definedName name="_xlnm._FilterDatabase" localSheetId="22" hidden="1">'RESIDENCIA SOTO'!$A$1:$AB$1</definedName>
```
*(Para `RESIDENCIA OVIEDO`, el nombre `_FilterDatabase` no está grabado en el XML base del libro, aplicándose dinámicamente mediante macro).*

##### B. Rango de Filtro en VBA ([BotonFiltrarSolicitudesxFechas.bas#L147-L151](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BotonFiltrarSolicitudesxFechas.bas#L147-L151))
```vba
' Aplicar filtro sólo en el rango real usado (A:AB según tu diseño)
ws.Range("A1:AB" & FILA_MAXIMA).AutoFilter Field:=2, _
    Criteria1:=">=" & criterioInicio, _
    Operator:=xlAnd, _
    Criteria2:="<" & criterioFin
```

##### C. Geometría de Columnas Ocultas en OpenXML (`<cols>`)
- **GIJÓN (`sheet2.xml`):**
  - Col 28 (`AB`): visible, width=47.
  - Col 29 (`AC`): visible, width=25.57.
  - Col 30 (`AD`, `REPARTIR COSTE`): visible, width=10.57.
  - Cols 31 a 16.384: **todas ocultas (`hidden="1"`)**, incluyendo `AJ` (36).
- **SOTO (`sheet23.xml`):**
  - Col 28 (`AB`): visible, width=50.
  - Col 29 (`AC`): visible, width=19.71.
  - Col 30 (`AD`, `REPARTIR COSTE`): visible, width=10.71.
  - Cols 31 a 16.384: **todas ocultas (`hidden="1"`)**, incluyendo `AJ` (36).
- **OVIEDO (`sheet34.xml`):**
  - Col 28 (`AB`): visible, width=12.71.
  - Cols 29 a 31 (`AC:AE`, funcionales): visibles, width=9.85.
  - Cols 32 y 33 (`AF, AG`, fórmulas): visibles.
  - Cols 34 a 16.384: **todas ocultas (`hidden="1"`)**, incluyendo `AJ` (36).

#### 2. Evaluación de Cobertura de Filtro y Riesgo de Ordenación Interactiva
- **Límite horizontal del AutoFilter:** En las tres residencias, el AutoFilter abarca estrictamente desde la columna **`A` hasta la columna `AB` (columnas 1 a 28)**.
- **Inclusión de columnas funcionales:**
  - GIJÓN (`AD` = 30): **EXCLUIDA** del rango de AutoFilter.
  - SOTO (`AD` = 30): **EXCLUIDA** del rango de AutoFilter.
  - OVIEDO (`AC, AD, AE` = 29, 30, 31): **EXCLUIDAS** del rango de AutoFilter.
- **Inclusión de la columna candidata AJ (36):** **EXCLUIDA** del rango de AutoFilter.
- **Inclusión de columnas ocultas:** El hecho de que una columna esté oculta (`hidden="1"`) en Excel **no hace que pertenezca al filtro**. Al estar definidas a la derecha de la columna 28 (`AB`), están físicamente fuera del rango `ref`.
- **Impacto sobre la Integridad en Ordenaciones Interactivas:**
  Si un usuario en Excel utiliza las flechas desplegables del encabezado del AutoFilter (por ejemplo, en columna A o B) para ordenar los datos ("Ordenar de A a Z"):
  - **Solo se reordenarán las celdas de las columnas `A:AB`**.
  - Las columnas funcionales (`AD` en GIJÓN/SOTO; `AC:AE` en OVIEDO) y la columna de identidad `AJ` **permanecerán fijas e inmóviles en sus filas físicas originales**.
  - **Consecuencia:** Se produciría una desalineación inmediata y catastrófica de los datos del huésped respecto a sus importes de repartir coste, tags y su futuro GUID de instancia.

---

### 3.3. Inserciones de Filas o Rangos

Se realizó una búsqueda global en todos los módulos VBA sobre instrucciones de inserción:
- `Rows.Insert`
- `EntireRow.Insert`
- `Insert Shift:=`

#### 1. Alcance de la Búsqueda y Coincidencias Localizadas
Se detectaron únicamente 3 coincidencias en todo el proyecto:
- [SeleccionHabitacionGIJON.frm#L278](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/SeleccionHabitacionGIJON.frm#L278): `wsBloq.Rows(j + 2).Insert Shift:=xlShiftDown`
- [SeleccionHabitacionOVIEDO.frm#L274](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/SeleccionHabitacionOVIEDO.frm#L274): `wsBloq.Rows(j + 2).Insert Shift:=xlShiftDown`
- [SeleccionHabitacionSOTO.frm#L294](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/SeleccionHabitacionSOTO.frm#L294): `wsBloq.Rows(j + 2).Insert Shift:=xlShiftDown`

#### 2. Hojas Afectadas
Estas tres instrucciones operan única y exclusivamente sobre las hojas auxiliares de bloqueo de habitaciones:
- `Habitaciones Bloqueadas GIJÓN`
- `Habitaciones Bloqueadas OVIEDO`
- `Apartamentos Bloqueados SOTO`

#### 3. Conclusión sobre Inserciones en Hojas Operativas
- **En las hojas `RESIDENCIA GIJÓN`, `RESIDENCIA SOTO` y `RESIDENCIA OVIEDO` no existe ninguna operación de inserción física de filas o rangos (`Rows.Insert` ni `Shift:=xlShiftDown`).**
- Las nuevas altas de expedientes nunca desplazan hacia abajo las filas existentes; se insertan agregando valores en la primera fila vacía al final de la tabla ([modDatabase.bas#L1734](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/modDatabase.bas#L1734)) o reconstruyendo la cuadrícula completa mediante `SincronizarHojaDesdeAccess`.

---

## 4. COMPROBACIÓN C — COBERTURA DE LIMPIEZA DE LA COLUMNA CANDIDATA AJ

Se ha analizado la cobertura horizontal de los bloques de vaciado en el procedimiento oficial [BorrarDatosAlojamientoOptimizado](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas#L21-L244) en [BorradoDatosAlojamiento.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/BorradoDatosAlojamiento.bas).

### 4.1. Verificación de los Bloques de Limpieza

#### 1. Bloques Específicos por Residencia (Líneas 178 a 195)
```vba
    Select Case ws.Name
        Case "RESIDENCIA GIJÓN"
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 20)).ClearContents
            ws.Range(ws.Cells(2, 22), ws.Cells(lastRow, 27)).ClearContents
            ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents
        Case "RESIDENCIA SOTO"
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 16)).ClearContents
            ws.Range(ws.Cells(2, 19), ws.Cells(lastRow, 20)).ClearContents
            ws.Range(ws.Cells(2, 22), ws.Cells(lastRow, 27)).ClearContents
            ws.Range(ws.Cells(2, 30), ws.Cells(lastRow, 30)).ClearContents
        Case "RESIDENCIA OVIEDO"
            ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, 21)).ClearContents
            ws.Range(ws.Cells(2, 23), ws.Cells(lastRow, 31)).ClearContents
    End Select
```

#### 2. Bloque de Saneamiento Automático de Columnas Excedentes (Líneas 116 a 124)
```vba
    Dim ultColUsada As Long, ultFilaUsada As Long
    ultColUsada = ws.UsedRange.Column + ws.UsedRange.Columns.Count - 1
    If ultColUsada > 36 Then
        ultFilaUsada = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If ultFilaUsada < lastRow Then ultFilaUsada = lastRow
        ws.Range(ws.Cells(1, 37), ws.Cells(ultFilaUsada, ultColUsada)).Clear
    End If
```

### 4.2. Respuestas a los Requerimientos Concretos

1. **¿Se borra actualmente AJ en GIJÓN?**
   **NO.** La última columna limpiada en GIJÓN es la 30 (`AD`). La columna 36 (`AJ`) no es tocada.
2. **¿Se borra actualmente AJ en SOTO?**
   **NO.** La última columna limpiada en SOTO es la 30 (`AD`). La columna 36 (`AJ`) no es tocada.
3. **¿Se borra actualmente AJ en OVIEDO?**
   **NO.** La última columna limpiada en OVIEDO es la 31 (`AE`). La columna 36 (`AJ`) no es tocada.
4. **¿El saneamiento automático desde AK afecta o no a AJ?**
   **NO AFECTA.** La condición es `If ultColUsada > 36` y el rango limpiado comienza explícitamente en `ws.Cells(1, 37)` (columna `AK`). La columna 36 (`AJ`) queda deliberadamente excluida del saneamiento superior.
5. **¿Existe algún otro bloque de limpieza en el procedimiento que alcance AJ?**
   **NO.** El único otro borrado es `ws.Range("P2:P" & lastRow).ClearComments` (columna 16).
6. **Dictamen de Cobertura sobre AJ:**
   Con la implementación vigente, **la columna candidata AJ (36) se encuentra en una zona ciega de limpieza**. Si se ubicara `IdInstanciaOrden` en `AJ`, **sobreviviría indebidamente a cualquier vaciado integral de alojamientos**, persistiendo identificadores obsoletos que contaminarían las altas posteriores.

---

## 5. COMPROBACIÓN D — MATRIZ FINAL DE COBERTURA

A continuación se consolida la matriz completa de operaciones que desplazan, ordenan, filtran o limpian las filas operativas:

| Operación | Procedimiento | Hoja | Rango Afectado | Última Columna Incluida | Tipo de Operación | AD / AC:AE Incluidos | AJ (36) Incluida | Resultado Funcional Acreditado |
| :--- | :--- | :--- | :--- | :---: | :--- | :---: | :---: | :--- |
| **Borrado por Orden** | `BorrarSolicitudesPorOrden` | GIJÓN | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AD`) | **SÍ** | Fila completa eliminada. Filas inferiores suben sincronizadas en todas las columnas. |
| **Borrado por Orden** | `BorrarSolicitudesPorOrden` | SOTO | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AD`) | **SÍ** | Fila completa eliminada. Filas inferiores suben sincronizadas en todas las columnas. |
| **Borrado por Orden** | `BorrarSolicitudesPorOrden` | OVIEDO | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AC:AE`) | **SÍ** | Fila completa eliminada. Filas inferiores suben sincronizadas en todas las columnas. |
| **Archivado Parcial** | `ArchivarOrdenesAntiguasBD` (Ruta 2) | GIJÓN | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AD`) | **SÍ** | Fila completa eliminada. Desplazamiento vertical homogéneo de toda la cuadrícula. |
| **Archivado Parcial** | `ArchivarOrdenesAntiguasBD` (Ruta 2) | SOTO | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AD`) | **SÍ** | Fila completa eliminada. Desplazamiento vertical homogéneo de toda la cuadrícula. |
| **Archivado Parcial** | `ArchivarOrdenesAntiguasBD` (Ruta 2) | OVIEDO | `ws.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** (`AC:AE`) | **SÍ** | Fila completa eliminada. Desplazamiento vertical homogéneo de toda la cuadrícula. |
| **Purga SICO-N2** | `PurgarOperacionesBorradasDeHojas` | GIJÓN, SOTO, OVIEDO | `wsBorr.Rows(fila)` | 16.384 (`XFD`) | Eliminación física (`Shift:=xlUp`) | **SÍ** | **SÍ** | Fila completa eliminada en purga offline. |
| **Archivado Total** | `ArchivarOrdenesAntiguasBD` (Ruta 1) | GIJÓN | `A2:T`, `V2:AA`, `AD2:AD` | 30 (`AD`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AD`) | **NO** | AJ sobrevive indebidamente al archivado total de la hoja. |
| **Archivado Total** | `ArchivarOrdenesAntiguasBD` (Ruta 1) | SOTO | `A2:P`, `S2:T`, `V2:AA`, `AD2:AD` | 30 (`AD`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AD`) | **NO** | AJ sobrevive indebidamente al archivado total de la hoja. |
| **Archivado Total** | `ArchivarOrdenesAntiguasBD` (Ruta 1) | OVIEDO | `A2:U`, `W2:AE` | 31 (`AE`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AC:AE`) | **NO** | AJ sobrevive indebidamente al archivado total de la hoja. |
| **Vaciado Integral** | `BorrarDatosAlojamientoOptimizado` | GIJÓN | `A2:T`, `V2:AA`, `AD2:AD` | 30 (`AD`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AD`) | **NO** | AJ sobrevive indebidamente al vaciado manual de la residencia. |
| **Vaciado Integral** | `BorrarDatosAlojamientoOptimizado` | SOTO | `A2:P`, `S2:T`, `V2:AA`, `AD2:AD` | 30 (`AD`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AD`) | **NO** | AJ sobrevive indebidamente al vaciado manual de la residencia. |
| **Vaciado Integral** | `BorrarDatosAlojamientoOptimizado` | OVIEDO | `A2:U`, `W2:AE` | 31 (`AE`) | Limpieza de celdas (`ClearContents`) | **SÍ** (`AC:AE`) | **NO** | AJ sobrevive indebidamente al vaciado manual de la residencia. |
| **Saneamiento Excedente** | `BorrarDatosAlojamientoOptimizado` | GIJÓN, SOTO, OVIEDO | `AK1:XFD<N>` | 16.384 (`XFD`) | Limpieza de celdas (`Clear`) | **NO** (empieza en 37) | **NO** (col 36 excluida) | Saneamiento limpia $\ge 37$ (`AK`), dejando `AJ` (36) intacta. |
| **Filtro por Fechas** | `MostrarFilasPorIntervaloFechas` | GIJÓN, SOTO, OVIEDO | `A1:AB<FILA_MAX>` | 28 (`AB`) | Filtro nativo (`AutoFilter`) | **NO** | **NO** | Solo oculta filas (`Hidden=True`), pero el control de filtro termina en `AB`. |
| **Filtro Nativo Libro** | `_FilterDatabase` (OpenXML) | GIJÓN, SOTO | `A1:AB1` | 28 (`AB`) | Definición de AutoFilter | **NO** | **NO** | Si el usuario ordena vía AutoFilter UI, desalinea `AD` y `AJ`. |
| **Ordenación VBA** | *(Ninguno)* | GIJÓN, SOTO, OVIEDO | *(No existe)* | — | Ordenación programada | — | — | **No existen ordenaciones VBA en las hojas operativas.** |
| **Inserción de Filas** | *(Ninguno)* | GIJÓN, SOTO, OVIEDO | *(No existe)* | — | Inserción con desplazamiento | — | — | **No existen inserciones VBA con desplazamiento en operativas.** |
| **Recarga Access** | `SincronizarHojaDesdeAccess` | GIJÓN, SOTO | Limpia `A:T/P`, `V:AA` | 27 (`AA`) | Limpieza y volcado matricial | **NO** | **NO** | Ni `AD` ni `AJ` se limpian ni se escriben; desalineación por *Row Drift*. |
| **Recarga Access** | `SincronizarHojaDesdeAccess` | OVIEDO | Limpia `A:U`, `W:AB` | 28 (`AB`) | Limpieza y volcado matricial | **NO** | **NO** | Ni `AC:AE` ni `AJ` se limpian ni se escriben; desalineación por *Row Drift*. |

---

## 6. CRITERIOS DE CONCLUSIÓN (HECHOS ACREDITADOS)

1. **Desplazamiento Conjunto en Eliminaciones:**
   Las operaciones vigentes que eliminan filas con desplazamiento vertical (`BorrarSolicitudesPorOrden`, `ArchivarOrdenesAntiguasBD` en su ruta parcial y `PurgarOperacionesBorradasDeHojas`) operan sobre el objeto nativo `ws.Rows(...)`. Por consiguiente, **eliminan físicamente la fila entera (columnas 1 a 16.384) y desplazan conjuntamente tanto las columnas funcionales (`AD` / `AC:AE`) como la columna candidata `AJ`**.
2. **Inexistencia de Borrados de Rangos Horizontales Parciales:**
   No se localizó en el código ninguna instrucción que ejecute un borrado parcial del tipo `Range("A2:AB2").Delete Shift:=xlUp`. Los borrados con desplazamiento son siempre de fila completa.
3. **Exclusión de Columnas Funcionales y AJ en el AutoFilter:**
   Tanto el rango persistido en el libro (`_FilterDatabase`) como la macro de filtrado Ribbon (`MostrarFilasPorIntervaloFechas`) limitan el AutoFilter a las columnas **`A:AB` (1 a 28)**.
   - Las columnas funcionales (`AD` en GIJÓN/SOTO, `AC:AE` en OVIEDO) quedan **fuera del AutoFilter**.
   - La columna candidata `AJ` (36) queda **fuera del AutoFilter**.
   - **Consecuencia:** Cualquier intento de ordenación interactiva ejecutado por un usuario a través del desplegable de AutoFilter en Excel provocaría la rotura inmediata de la correlación entre el expediente y sus datos funcionales o su identidad.
4. **Ausencia de Inserciones con Desplazamiento en Hojas Operativas:**
   En todo el proyecto no existe ninguna llamada a `Rows.Insert` o `Insert Shift:=` sobre las hojas de residencia. Las altas se agregan al final de la tabla o se reconstruyen secuencialmente desde Access.
5. **Zona Ciega de Limpieza en la Columna Candidata AJ:**
   La columna `AJ` (36) **no se borra** en `BorrarDatosAlojamientoOptimizado` ni en la ruta de vaciado total de `ArchivarOrdenesAntiguasBD`. Al comenzar el saneamiento de columnas excedentes en la columna 37 (`AK`), la columna `AJ` sobrevive indemne a los vaciados integrales, lo que constituye una limitación física de cobertura que debe resolverse si se aprueba dicha columna.
6. **Desconexión en la Descarga desde Access (`SincronizarHojaDesdeAccess`):**
   La reconstrucción desde Access limpia y escribe únicamente hasta la columna `AA` (27) en GIJÓN/SOTO y `AB` (28) en OVIEDO. Ni las columnas funcionales ni `AJ` son gestionadas durante la sincronización, confirmándose el riesgo de *Row Drift* acreditado en la Orden 1W.

---

## 7. RELACIÓN DE SCRIPTS AUXILIARES UTILIZADOS

En cumplimiento de las restricciones obligatorias, no se introdujo ningún script en el repositorio de trabajo. Todas las herramientas de inspección se ejecutaron en el directorio scratch externo:

- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\search_deletions.py` (auditoría de borrados).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\check_entire_row_delete.py` (verificación de borrados de fila completa).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\search_sorts.py` y `search_wider_sort.py` (búsqueda de operaciones de ordenación).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\inspect_openxml.py` (inspección de OpenXML y columnas ocultas).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\check_defined_names.py` y `print_all_defined_names.py` (inspección de `_FilterDatabase`).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\search_autofilter_vba.py` y `read_filtrar_fechas.py` (análisis del filtro por fechas).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\inspect_headers.py` (extracción de encabezados fila 1).
- `C:\Users\jaime_ntymfau\.gemini\antigravity\brain\1bda4f34-55a5-4d42-b9bf-6092244693f2\scratch\search_inserts.py` (búsqueda de operaciones de inserción).

---

## 8. ESTADO FINAL DE GIT

- **Comando:** `git status --short`
- **Resultado:** Idéntico al estado inicial previo a la orden. **Cero modificaciones en archivos del repositorio, libro Excel o base de datos Access.**
