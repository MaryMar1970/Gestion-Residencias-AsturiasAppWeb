# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: 2026-10-27 01:10 (Solución Integral de Solapamientos de Reservas y Avisos Emergentes en Calendarios Gijón, Oviedo y Soto)

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  *Tamaño:* 6.905.800 bytes | *MD5:* `802BE59E8C5E49986BE4071CDE3AB2DE` | *SHA-256:* `66A099850971439159B1428BEBE22EA5732BD41B14E52385465C03D8229D496F`
- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
  *Tamaño:* 6.742.016 bytes | *MD5:* `080134A9F53A6D41060F69A6B907214B` | *SHA-256:* `B5418FA3EBA18B4B84B8C4B70171FD937C4142A4C594586FDBE8A4DE552DAD38`
- **Copia Activa de Trabajo:**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`
- **Copia Desplegada en Red (Producción):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`

---

### 2. Inventario de Componentes Clave Modificados en este Hito

1. `ModuloCalendarioOviedo.bas`:
   - Incorporación de guardas defensivas de color en `RepintarReservasPagadas` en celdas de check-in, estancia intermedia y check-out:
     `If wsCal.Cells(...).Interior.color <> vbRed And wsCal.Cells(...).Interior.color <> COLOR_SOLAPE_ACEPTADO Then`
   - Impide que el repintado de reservas pagadas sobreescriba celdas rojas de solape o solapes aceptados.

2. `ModuloCalendarioGijon.bas`:
   - Restitución en el flujo principal de `ActualizarCalendarioInterno` de la llamada interactiva de notificación:
     ```vba
     If huboSolape Then
         MostrarDetalleSolapes
     End If
     ```
   - Incorporación de guardas defensivas de color en `RepintarReservasPagadas` para celdas con `vbRed` y `COLOR_SOLAPE_ACEPTADO`.

3. `ModuloCalendarioSoto.bas`:
   - Incorporación de guardas defensivas de color en `RepintarReservasPagadas` para celdas con `vbRed` (check-in y días intermedios).

---

### 3. Preservación y Compatibilidad Integral
- **Total de componentes VBA:** 171 componentes preservados al 100%.
- **Errores de compilación:** 0 errores (`VBE.CommandBars.FindControl(1, 578).Enabled == False`).
- **Codificación:** 100% Windows-1252 ANSI, Cero Mojibake verificado con guardián automatizado.
- **Regla estricta de nombres:** Texto literal `GIJÓN` respetado sin `Chr(211)`.

---

### 4. Procedimiento de Restauración Inmediata
Si en futuras modificaciones se requiriese restaurar este estado de referencia consolidado:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
