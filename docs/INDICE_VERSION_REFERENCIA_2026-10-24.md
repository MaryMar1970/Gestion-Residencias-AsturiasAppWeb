# ÍNDICE DE VERSIÓN DE REFERENCIA OFICIAL
## Hito: 2026-10-24 00:45 (Fases 1 y 2 Consolidadas)

### 1. Estado y Ubicación de las Réplicas Oficiales

- **Libro Front-End Oficial (Protegido Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\BDAS_Multiusuario_VERSION_REFERENCIA.xlsm`
  *Tamaño:* 6.838.715 bytes | *MD5:* `C61E591EDD8BCAB9D789949D471FF617` | *SHA-256:* `FDDF25F41FE8C11B3FB6D0F3B4756CA1473E85E3250D20227FB034FD2C75C7F6`
- **Base de Datos Back-End Oficial (Protegida Solo Lectura):**
  `H:\ResidenciaBD\Plantilla\VersionReferencia\Residencia_BE_VERSION_REFERENCIA.accdb`
  *Tamaño:* 6.733.824 bytes | *MD5:* `A5B86926058CCF73DEDA7F0C8F83A323` | *SHA-256:* `DDA1D01BFE892B551BFBA2E414D8F487473C32392BFAD195ACA69903131D6A89`
- **Copia Activa de Trabajo:**
  `H:\ResidenciaApp\bdas-multiusuario\BDAS_Multiusuario.xlsm`
- **Copia Desplegada en Red (Producción):**
  `H:\ResidenciaBD\BDAS_Multiusuario.xlsm`

---

### 2. Inventario de Componentes Clave Modificados en este Hito

1. `ThisWorkbook.cls`:
   - Secuencia de inicialización en `Workbook_Open` libre de parpadeos con consolidación de buffer gráfico.
   - Idempotencia geométrica en `AjustarCaratulaInicio` ($\pm 2\text{ pt}$).
   - Reposición automática de barras de desplazamiento, encabezados y cuadrículas en hojas operativas (`Workbook_SheetActivate`).
2. `modDatabase.bas`:
   - Pre-carga en memoria (`Load frmLogin`) antes de mostrar diálogo.
   - Congelación de repintado (`Application.screenUpdating = False`) antes de descargar login.
   - Pre-carga de `frmSelectorResidencia` y reactivación en el instante de mostrarlo.
   - Preservación íntegra del bucle de intentos (`MAX_INTENTOS_LOGIN = 3`) y de todos los flujos de activación y auditoría.
   - Refactorización de persistencia y consultas SQL unificadas (REF-01 a REF-16).
3. `frmLogin.frm`:
   - Tarjeta Fluent Card estilizada (fondo blanco puro, barra de acento verde Guardia Civil `RGB(27, 67, 50)` de 3.5 pt).
   - Inputs planos con borde dinámico en foco.
   - Sincronización de tecla `Enter` consumiendo evento (`KeyCode = 0`).
   - Supresión del reposicionamiento en `Activate` (erradicación de parpadeo de bandas).
4. `frmSelectorResidencia.frm`:
   - Tarjeta Fluent Card con cabecera verde institucional y lista de selección plana.
   - Normalización de nombres de sedes (`GIJÓN`, `OVIEDO`, `SOTO DEL BARCO`).
5. `frmAviso.frm`:
   - Notificación flotante Toast no modal (`Show vbModeless`) estilizada con fondo blanco y franja verde superior.
   - Desprogramación defensiva del timer previo de `Application.OnTime`.
   - Preservación de controles nativos `lblAviso` y métodos públicos `MostrarAviso` y `CerrarAviso`.
6. `ModuloLOG.bas`:
   - Saneamiento de auditoría en hojas LOG y protección de inyección de fórmulas.
7. Módulos y Hojas de Facturación y Resumen:
   - Cuadre contable exacto al céntimo en facturas PAX (`AsignarNumFactura.bas`, `MarcarSiPagadosEnResidencia.bas`, `ReevaluacionSolicitudes.bas`, `Hoja9`, `Hoja7`).

---

### 3. Procedimiento de Restauración Inmediata
Si en futuras modificaciones se requiriese restaurar este estado de referencia consolidado:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
