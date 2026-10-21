# Índice Técnico de la Versión de Referencia Oficial — Hito 2026-10-21 23:30

> **Ubicación de Referencia Protegida:** `H:\ResidenciaBD\Plantilla\VersionReferencia\`  
> **Libro Maestro:** `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` (MD5: `2507104C574B3DD48E11DD9A0FD76A76` | 7.188.546 bytes)  
> **Base de Datos Maestro:** `Residencia_BE_VERSION_REFERENCIA.accdb` (MD5: `D51AE318D49BC74C4BC738F5B58B2B51` | 6.725.632 bytes)  
> **Estado de Atributos:** Ambos archivos protegidos contra escritura (`IsReadOnly = True`).  
> **Script de Recuperación:** `bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1`  

---

## 1. Alcance Funcional y Técnico Consolidado en este Hito

Esta versión de referencia consolida la totalidad de mejoras del hito previo (2026-10-21 00:55) e incorpora las siguientes soluciones y blindajes técnicos:

1. **Auto-Fit Dinámico Proporcional de la Carátula (`ThisWorkbook.cls`)**:
   - Detección en tiempo real del área cliente de Excel (`Application.UsableWidth` y `Application.UsableHeight`).
   - Reescalado automático proporcional de la forma gráfica `ImgCaratulaCasaCuartel` respetando la relación de aspecto nativa $1.858 : 1$ ($1570 \times 845\text{ px}$).
   - Centrado horizontal y vertical en pantalla, asegurando que el tercio inferior (calzada empedrada y base) se visualice íntegro y sin recortes en cualquier monitor (portátiles 1080p al 125%, sobremesa o pantallas 4K).
   - Manejo seguro de desprotección y reprotección mediante `ModuloConfigSegura.ObtenerPasswordHojas()`.

2. **Tarjeta de Login Estilizada y Anclada Milimétricamente (`frmLogin.frm`)**:
   - Rediseño compacto tipo tarjeta ($240\text{ pt}$ ancho $\times 165\text{ pt}$ alto) con fondo blanco puro y botón de entrada en verde institucional sobrio (`RGB(35, 75, 45)`).
   - Posicionamiento manual exacto (`StartUpPosition = 0`) calculado en `PosicionarFormulario`: centrado en el eje horizontal y situado al $52\%$ de la altura útil, encajado con precisión milimétrica en el patio empedrado entre los dos guardias civiles.
   - Ambos guardias civiles (a izquierda y derecha) quedan completamente despejados y visibles.
   - Navegación optimizada por teclado: foco automático inicial en Usuario, salto con `Enter` a Clave, y ejecución inmediata de `ProcesarLogin` con `Enter` desde la casilla de clave.

3. **Erradicación del Reloj de Arena en Columna T de `Hoja2 (RESIDENCIA GIJÓN)`**:
   - Supresión de la condición redundante y prematura `Or Target.Column = 20` dentro de `Worksheet_SelectionChange`.
   - Eliminación de la ejecución parásita de `ValidateHabitaciones` al hacer clic o desplazarse por la columna *DTO. FAM. NUM.*
   - Se mantiene intacta la validación de coherencia habitacional en la columna **AB** (*Grabación Solicitud*), equiparando el comportamiento fluido de Gijón con el de Soto (columna AB) y Oviedo (columna AF).

4. **UX de Cierre Suave y Confirmación Interactiva de Salida (`ThisWorkbook.cls`)**:
   - Diálogo interactivo ante el evento de cierre por aspa: `MsgBox("¿Desea salir de la aplicación?", vbQuestion + vbYesNo + vbDefaultButton2)`.
   - Transición fluida y elegante con bloqueo de parpadeo (`Application.ScreenUpdating = False`), minimización estética y guardado ordenado.

5. **Gobernanza Técnica: Protocolo Estricto de Codificación y Cero Mojibake en VBA**:
   - Blindaje de 3 capas en `.agents/AGENTS.md` (ANSI Windows-1252 nativo, prohibición de VBScript para textos en español y control de puerta pre-entrega).

6. **Saneamiento Ortográfico Integral en Módulo de Control de Duplicados (`EvitarDuplicidadSolicitudesGyS.bas`)**:
   - Sustitución de caracteres de flechas desalineados `âž”` por `->` y acentuación limpia Windows-1252 ANSI.

7. **Tarifas Dinámicas en Facturas PAX (`FACTURA GIJÓN` y `FACTURA OVIEDO`)**:
   - Fórmulas dinámicas matriciales con `SUMAPRODUCTO` evaluando la fecha de estancia (`F19`) contra los tramos temporales en `TARIFAS UNIDAD`.

8. **Saneamiento Estructural de Hojas de Auditoría LOG (`LOG_GIJÓN`, `LOG_SOTO`, `LOG_OVIEDO`)**:
   - Parseo estricto por columnas: `[Usuario, Fecha/Hora, Nº Orden, Columna, Valor anterior, Valor nuevo]`.
   - Protección contra inyección de fórmulas (`'=`).
   - Sincronización relacional inversa de los 500 registros más recientes y desambiguación de eventos `LOGIN` por residencia.

9. **Sistema de Sesión Persistente y Auto-reparable**:
   - Persistencia de credenciales en Nombres Definidos ocultos de Excel inmune a la volatilidad de memoria VBA.
   - Erradicación del Error 50290 en `Worksheet_Calculate` de `Hoja11` (`SOBRE GIJÓN`).
   - Blindaje permanente de la barra de estado.

10. **Preservación Total e Incondicional de los 171 Componentes VBA**:
    - Todos los componentes y hojas se mantienen operativos con 0 errores de compilación.

---

## 2. Histórico de Versiones de Referencia Archivadas

Ubicación: `H:\ResidenciaBD\Plantilla\VersionReferencia\Historico\`

| Hito | Archivo XLSM | Archivo ACCDB | MD5 XLSM | MD5 ACCDB | Notas Principales |
| :---: | :---| :---| :---:| :---:| :---|
| **21/10/2026 23:30** | `BDAS_Multiusuario_VERSION_REFERENCIA.xlsm` *(Activa)* | `Residencia_BE_VERSION_REFERENCIA.accdb` *(Activa)* | `2507104C574B3DD48E11DD9A0FD76A76` | `D51AE318D49BC74C4BC738F5B58B2B51` | Auto-Fit Carátula, Tarjeta Login entre Guardias, Erradicación Reloj de Arena Col T Gijón |
| **21/10/2026 00:55** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-21_0055.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-21_0055.accdb` | `E8AE2E58DC520C55042E93C8992945F9` | `E3586759300882BCEEEB1677F31643B7` | Cierre Suave UX, Protocolo Cero Mojibake, Duplicados flechas y ortografía, Tarifas Dinámicas |
| **20/10/2026 23:15** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_2315.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_2315.accdb` | `5A23A4AB3A34A88442E6C28D5B1FB021` | `2D5E9D788825B17AAECEF46531D9ADFC` | Consolidación 171 componentes, Contador facturas, Saneamiento LOGs, Sesión persistente |
| **20/10/2026 19:35** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_1935.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_1935.accdb` | `135FDB2285F4983E232D59C76C0074E2` | `C40AA4331F3B83ACACD5132707ED6BAA` | Tarifas dinámicas Facturas Gijón/Oviedo, Saneamiento LOGs, Fix 50290 |
| **20/10/2026 01:00** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-20_0100.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-20_0100.accdb` | `1B7B1950F5071D2927ED31BEBC7B2F5D` | `54522630B0BB8D2229EDA29029E22BFD` | Auditoría LOGs, Sincronización Top 500, Fix 50290, Sesión y Rol persistente |
| **19/10/2026 21:35** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_2135.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_2135.accdb` | `57821D0835BCBB82B5356A722B982945` | `DD7C3F0DC0D62B4DAEF487DE9D2F303D` | Anti-Mojibake 5 capas, Normaliz. ASCII, Carátula, Reservas UPSERT |
| **19/10/2026 01:50** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-19_0150.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-19_0150.accdb` | `54D94F60F542DE5210795C3C30A5A419` | `EDE9C7ECB540604F835FACC39F8F08B8` | Carátula Casa Cuartel, Turno Soto, Persistencia ADODB |
| **18/10/2026 21:25** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_2125.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_2125.accdb` | `02B9E7FEADCE6632CB63445E4860B777` | `5C84D4B29BB88F7C9E0FF9DF515E34F6` | Formulario Lista Negra, RibbonX roles, Purga LOGs |
| **18/10/2026 03:20** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-18_0320.xlsm` | `Residencia_BE_VERSION_REFERENCIA_2026-10-18_0320.accdb` | `5F0532FF42BA67C7E9BBCB2DF578DA2D` | `6426624 bytes` | Fase 1 Roles centralizados |
| **17/10/2026 23:15** | `BDAS_Multiusuario_VERSION_REFERENCIA_2026-10-17_2315.xlsm` | — | `9E9685D7F5F8E33B8A139EE4929AE543` | — | Tarifas temporales, blindaje de fórmulas |

---

## 3. Protocolo de Reversión

En caso de cualquier eventualidad o necesidad de recuperar este estado estable:
```powershell
powershell -ExecutionPolicy Bypass -File "H:\ResidenciaApp\bdas-multiusuario\scripts\Restaurar-Version-Referencia.ps1"
```
El script generará una copia de seguridad preventiva del estado de trabajo y restaurará de forma íntegra e instantánea los archivos maestro protegidos.
