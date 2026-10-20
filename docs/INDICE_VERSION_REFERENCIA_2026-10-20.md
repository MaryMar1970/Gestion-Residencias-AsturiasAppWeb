# Índice Técnico de la Versión de Referencia Oficial — Hito 2026-10-20 19:35

> **Ubicación de Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`  
> **Libro Maestro:** `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (MD5: `135fdb2285f4983e232d59c76c0074e2` | 7.152.489 bytes)  
> **Base de Datos Maestro:** `Residencia_BE_VERSION_REFERENCIA.accdb` (MD5: `c40aa4331f3b83acacd5132707ed6baa` | 6.721.536 bytes)  
> **Estado de Atributos:** Ambos archivos protegidos contra escritura (`IsReadOnly = True`).  
> **Script de Recuperación:** `bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1`  

---

## 1. Alcance Funcional y Técnico Consolidado en este Hito

Esta versión de referencia consolida la totalidad de mejoras del hito previo (2026-10-20 01:00) e incorpora la resolución definitiva de tarifas dinámicas en facturación:

1. **Tarifas Dinámicas en Facturas PAX (`FACTURA GIJÓN` y `FACTURA OVIEDO`)**:
   - Desvinculación de referencias estáticas fijas en las celdas de precio/día de habitación individual (`R23`), doble (`R24`) y cama supletoria (`R25`).
   - Implementación de fórmulas matriciales dinámicas con `SUMAPRODUCTO` evaluando la fecha de estancia (`F19`, equivalente a `L2` de las hojas de residencia) contra los tramos temporales vigentes en `TARIFAS UNIDAD`:
     - Gijón: Rango `$E$3:$O$9`.
     - Oviedo: Rango `$E$13:$O$19`.
   - Cuadre contable automático al 100% al céntimo entre la solicitud de residencia (Columna U en Gijón, Columna V en Oviedo) y la factura emitida, garantizando total autonomía en futuros ejercicios sin requerir mantenimiento manual de fórmulas.

2. **Saneamiento Estructural de Hojas de Auditoría LOG (`LOG_GIJÓN`, `LOG_SOTO`, `LOG_OVIEDO`)**:
   - Corrección del algoritmo de despiece y parseo de cadenas delimitadas en [ModuloLOG.bas](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/ModuloLOG.bas):
     - `Columna A`: Usuario.
     - `Columna B`: Fecha y Hora (`dd/mm/yyyy hh:mm`).
     - `Columna C`: Nº de Orden.
     - `Columna D`: Columna afectada o acción (`LOGIN` / `CAMBIO (...)`).
     - `Columna E`: Valor anterior limpio.
     - `Columna F`: Valor nuevo limpio.
   - Erradicación de la sobreescritura de campos concatenados en columna F.

3. **Sincronización Relacional Invertida de los 500 Registros Más Recientes**:
   - Corrección de la consulta SQL en `SincronizarHojaLOGDesdeAccess`: sustitución de `SELECT TOP 500 ... ORDER BY Id ASC` por subconsulta anidada descendente:
     ```sql
     SELECT * FROM (SELECT TOP 500 * FROM LogActividad WHERE ... ORDER BY Id DESC) ORDER BY Id ASC
     ```
   - Carga inmediata y fiel de la actividad más reciente en las tres residencias.

4. **Filtrado Estricto de Eventos de LOGIN por Residencia**:
   - Desambiguación de cláusulas SQL para eventos de acceso: los logins se filtran estrictamente por la residencia correspondiente (`Accion = 'LOGIN' AND Detalle LIKE '%RESIDENCIA%'`), evitando contaminación cruzada entre hojas LOG.

5. **Blindaje contra Inyección de Fórmulas en Celdas de Auditoría**:
   - Detección preventiva de valores que comienzan por `=` o `+` e inserción con prefijo de apóstrofo (`'`), erradicando el error en tiempo de ejecución 1004.

6. **Erradicación del Error 50290 en `Worksheet_Calculate` de `Hoja11` (`SOBRE GIJÓN`)**:
   - Eliminación de la manipulación de `Application.EnableEvents = False/True` dentro del evento de cálculo activo al interactuar con desplegables (como columna F *Situación*).

7. **Sistema de Sesión Persistente y Auto-reparable (Inmune a Volatilidad VBA)**:
   - Persistencia de credenciales (`m_NombreUsuario`, `m_Rol`, `m_ResidenciaActiva`) en Nombres Definidos ocultos de Excel (`_Session_Usuario`, `_Session_Rol`, `_Session_Residencia`) y rehidratación automática.

8. **Blindaje Permanente de la Barra de Estado**:
   - Preservación permanente de la leyenda de conexión en la barra inferior sin que sea sobrescrita por alertas de accesibilidad de Office.

---

## 2. Histórico de Versiones de Referencia Archivadas

Ubicación: `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`

| Hito | Archivo XLSM | Archivo ACCDB | MD5 XLSM | MD5 ACCDB | Notas Principales |
| :---: | :---| :---| :---:| :---:| :---|
| **20/10/2026 19:35** | `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` *(Activa)* | `Residencia_BE_VERSION_REFERENCIA.accdb` *(Activa)* | `135fdb2285f4983e232d59c76c0074e2` | `c40aa4331f3b83acacd5132707ed6baa` | Tarifas dinámicas Facturas Gijón/Oviedo, Saneamiento LOGs, Fix 50290, Sesión persistente |
| **20/10/2026 01:00** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_0100.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_0100.accdb` | `1b7b1950f5071d2927ed31bebc7b2f5d` | `54522630b0bb8d2229eda29029e22bfd` | Auditoría LOGs, Sincronización Top 500, Fix 50290, Sesión y Rol persistente |
| **19/10/2026 21:35** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_2135.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_2135.accdb` | `57821d0835bcbb82b5356a722b982945` | `dd7c3f0dc0d62b4daef487de9d2f303d` | Anti-Mojibake 5 capas, Normaliz. ASCII, Carátula, Reservas UPSERT |
| **19/10/2026 01:50** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_0150.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_0150.accdb` | `54d94f60f542de5210795c3c30a5a419` | `ede9c7ecb540604f835facc39f8f08b8` | Carátula Casa Cuartel, Turno Soto, Persistencia ADODB |
| **18/10/2026 21:25** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_2125.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_2125.accdb` | `02b9e7feadce6632cb63445e4860b777` | `5c84d4b29bb88f7c9e0ff9df515e34f6` | Formulario Lista Negra, RibbonX roles, Purga LOGs |
| **18/10/2026 03:20** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_0320.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_0320.accdb` | `5f0532ff42ba67c7e9bbcb2df578da2d` | `6426624 bytes` | Fase 1 Roles centralizados |
| **17/10/2026 23:15** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-17_2315.xlsm` | — | `9e9685d7f5f8e33b8a139ee4929ae543` | — | Tarifas temporales, blindaje de fórmulas |

---

## 3. Protocolo de Reversión

En caso de cualquier eventualidad o necesidad de recuperar este estado estable:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
El script generará una copia de seguridad preventiva del estado de trabajo y restaurará de forma íntegra e instantánea los archivos maestro protegidos.