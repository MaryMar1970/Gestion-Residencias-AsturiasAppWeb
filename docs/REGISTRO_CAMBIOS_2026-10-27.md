# REGISTRO DE CAMBIOS - 2026-10-27

## Hito 2026-10-27 01:10: Consolidación y Promoción a Versión de Referencia Oficial (Solución Integral de Solapamientos de Reservas y Avisos Emergentes en Calendarios Gijón, Oviedo y Soto)

### 1. Resumen Ejecutivo
Se ha guardado y promocionado formalmente la versión activa de trabajo `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` a **Versión de Referencia Oficial**, estableciendo las réplicas protegidas contra escritura en el repositorio central de plantillas maestras (`H:\ResidenciaBD\Plantilla\VersionReferencia\`), archivando el hito previo (`2026-10-25 23:45`) en el histórico (`Historico\`), actualizando los scripts de promoción y restauración inmediata (`Promocionar-Version-Referencia.ps1`, `Restaurar-Version-Referencia.ps1`), registrando exhaustivamente la trazabilidad de copias de seguridad incrementales diarias vinculadas al código y sincronizando la copia de despliegue en red (`H:\ResidenciaBD\BDAS_Multiusuario.xlsm`).

En esta versión se han resuelto de forma definitiva las anomalías en la detección, visualización gráfica y notificación interactiva de solapamientos de reservas en las tres sedes del sistema:

> **1. Preservación Incondicional del Pintado en Rojo (`vbRed`) de Celdas Solapadas:**
> 
> En los tres calendarios operativos (`Calendario OVIEDO`, `Calendario GIJÓN` y `Calendario SOTO`), cuando se registraba un solapamiento entre reservas o entre reservas y bloqueos, el algoritmo principal detectaba el conflicto y marcaba correctamente las celdas afectadas con fondo rojo (`vbRed`). Sin embargo, en la fase final de actualización del calendario, la subrutina `RepintarReservasPagadas` recorría las reservas con estado de pago "SÍ" o parcial y aplicaba el color verde claro (`RGB(198, 239, 206)`) sobre todas sus celdas sin evaluar si alguna de ellas estaba en solape, tapando el color rojo de alerta.
> 
> **Solución aplicada:**
> Se incorporaron guardas lógicas condicionales en los puntos de pintado de `RepintarReservasPagadas` en los tres módulos de calendario (`ModuloCalendarioOviedo.bas`, `ModuloCalendarioGijon.bas` y `ModuloCalendarioSoto.bas`). Ahora, antes de aplicar el color de pagado sobre una celda (tanto en check-in, estancia intermedia o check-out), el procedimiento valida que la celda no contenga `vbRed` (ni `COLOR_SOLAPE_ACEPTADO` en Oviedo y Gijón), garantizando que las advertencias visuales de solapamiento permanezcan visibles en todo momento.

> **2. Restitución del Diálogo Flotante de Aviso en Calendario GIJÓN (`MostrarDetalleSolapes`):**
> 
> En la sede de Gijón existía una discrepancia funcional: aunque el motor calculaba los solapes (`huboSolape = True`) y pintaba las celdas, nunca se presentaba la ventana emergente con el desglose detallado de apartamentos, números de orden en conflicto y fechas afectadas.
> 
> **Solución aplicada:**
> Se identificó la ausencia de la llamada en el flujo principal de `ModuloCalendarioGijon.ActualizarCalendarioInterno` y se restituyó de inmediato el bloque:
> ```vba
> If huboSolape Then
>     MostrarDetalleSolapes
> End If
> ```
> con lo que el comportamiento de notificación interactiva queda completamente homogeneizado y operativo en las tres residencias.

---

### 2. Identificadores Forenses y Verificación de Integridad

| Archivo / Réplica | Ubicación | Tamaño (Bytes) | Hash MD5 | Hash SHA-256 | Atributos |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Versión de Referencia Oficial (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` | 6.905.800 | `802BE59E8C5E49986BE4071CDE3AB2DE` | `66A099850971439159B1428BEBE22EA5732BD41B14E52385465C03D8229D496F` | Solo Lectura (`ReadOnly = True`) |
| **Versión de Referencia Oficial (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb` | 6.742.016 | `080134A9F53A6D41060F69A6B907214B` | `B5418FA3EBA18B4B84B8C4B70171FD937C4142A4C594586FDBE8A4DE552DAD38` | Solo Lectura (`ReadOnly = True`) |
| **Versión Activa de Trabajo (App)** | `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm` | 6.905.800 | `802BE59E8C5E49986BE4071CDE3AB2DE` | `66A099850971439159B1428BEBE22EA5732BD41B14E52385465C03D8229D496F` | Lectura / Escritura |
| **Copia Desplegada en Producción** | `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` | 6.905.800 | `802BE59E8C5E49986BE4071CDE3AB2DE` | `66A099850971439159B1428BEBE22EA5732BD41B14E52385465C03D8229D496F` | Lectura / Escritura |
| **Base de Datos Activa (Backend)** | `H:\ResidenciaBD\Residencia_BE.accdb` | 6.742.016 | `080134A9F53A6D41060F69A6B907214B` | `B5418FA3EBA18B4B84B8C4B70171FD937C4142A4C594586FDBE8A4DE552DAD38` | Lectura / Escritura |
| **Copia Cierre Ref. XLSM (01:10)** | `H:\ResidenciaBD\BackupsDiarios\BDAS_Multiusuario_BACKUP_2026-10-27_0110_VERSION_REFERENCIA.xlsm` | 6.905.800 | `802BE59E8C5E49986BE4071CDE3AB2DE` | `66A099850971439159B1428BEBE22EA5732BD41B14E52385465C03D8229D496F` | Respaldo Incremental |
| **Copia Cierre Ref. ACCDB (01:10)** | `H:\ResidenciaBD\BackupsDiarios\Residencia_BE_BACKUP_2026-10-27_0110_VERSION_REFERENCIA.accdb` | 6.742.016 | `080134A9F53A6D41060F69A6B907214B` | `B5418FA3EBA18B4B84B8C4B70171FD937C4142A4C594586FDBE8A4DE552DAD38` | Respaldo Incremental |
| **Histórico Ref. 25-Oct (XLSM)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-25_2345.xlsm` | 6.832.371 | `9A241B718DC68AD827A93F63A50384C1` | `54B580B1C2777617195E84C257D7FD70771C50B5C518948AC5728B308E528590` | Solo Lectura (`ReadOnly = True`) |
| **Histórico Ref. 25-Oct (ACCDB)** | `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\Residencia_BE_VERSION_REFERENCIA_2026-10-25_2345.accdb` | 6.737.920 | `5C69DAFA82711161C0F2C2D023689A53` | `A253096DEF96D646710D2113057905BE75B6CAD6D0B2C2EC4DE4CD729F707406` | Solo Lectura (`ReadOnly = True`) |

---

### 3. Registro Detallado de Modificaciones en Componentes

#### A. Módulo `ModuloCalendarioOviedo.bas`
- Procedimiento `RepintarReservasPagadas`:
  * En pintado de entrada / check-in (tarde):
    ```vba
    If wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color <> vbRed And _
       wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color <> COLOR_SOLAPE_ACEPTADO Then
        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = colorPagado
    End If
    ```
  * En pintado de días completos intermedios (mañana, tarde y noche):
    ```vba
    If wsCal.Cells(reservaArray(0), col + tramo).Interior.color <> vbRed And _
       wsCal.Cells(reservaArray(0), col + tramo).Interior.color <> COLOR_SOLAPE_ACEPTADO Then
        wsCal.Cells(reservaArray(0), col + tramo).Interior.color = colorPagado
    End If
    ```
  * En pintado de salida / check-out (mañana):
    ```vba
    If wsCal.Cells(reservaArray(0), colCheckout).Interior.color <> vbRed And _
       wsCal.Cells(reservaArray(0), colCheckout).Interior.color <> COLOR_SOLAPE_ACEPTADO Then
        wsCal.Cells(reservaArray(0), colCheckout).Interior.color = colorPagado
    End If
    ```

#### B. Módulo `ModuloCalendarioGijon.bas`
- Procedimiento `ActualizarCalendarioInterno`:
  * Restitución de la llamada de aviso interactivo al finalizar la fase de procesamiento y pintado:
    ```vba
    If huboSolape Then
        MostrarDetalleSolapes
    End If
    ```
- Procedimiento `RepintarReservasPagadas`:
  * En pintado de check-in (tarde):
    ```vba
    If wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color <> vbRed And _
       wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color <> COLOR_SOLAPE_ACEPTADO Then
        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = colorPagado
    End If
    ```
  * En pintado de días intermedios (mañana, tarde y noche):
    ```vba
    If wsCal.Cells(reservaArray(0), col + tramo).Interior.color <> vbRed And _
       wsCal.Cells(reservaArray(0), col + tramo).Interior.color <> COLOR_SOLAPE_ACEPTADO Then
        wsCal.Cells(reservaArray(0), col + tramo).Interior.color = colorPagado
    End If
    ```

#### C. Módulo `ModuloCalendarioSoto.bas`
- Procedimiento `RepintarReservasPagadas`:
  * En pintado de check-in (tarde):
    ```vba
    If wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color <> vbRed Then
        wsCal.Cells(reservaArray(0), colCheckin + 2).Interior.color = colorPagado
    End If
    ```
  * En pintado de días intermedios (mañana, tarde y noche):
    ```vba
    If wsCal.Cells(reservaArray(0), col + tramo).Interior.color <> vbRed Then
        wsCal.Cells(reservaArray(0), col + tramo).Interior.color = colorPagado
    End If
    ```

---

### 4. Verificación de Cumplimiento de Reglas del Proyecto
- **Preservación de Componentes:** 171 componentes en el proyecto VBA (100% preservados).
- **Compilación en Excel COM:** 0 errores de compilación (`VBE.CommandBars.FindControl(1, 578).Enabled == False`).
- **Codificación:** 100% Windows-1252 ANSI en todos los módulos exportados a `vba-modules/` (`Validar-Codificacion-VBA.py` -> 0 problemas detectados).
- **Regla Estricta de Nombres:** Literal `GIJÓN` verificado en todos los componentes modificados sin uso de `Chr(211)`.
- **Inmutabilidad del Golden Master:** La réplica Golden Master permanece estrictamente inmutable e intacta (`BDAS_Multiusuario_GOLDEN_MASTER.xlsm`, `ReadOnly = True`).
