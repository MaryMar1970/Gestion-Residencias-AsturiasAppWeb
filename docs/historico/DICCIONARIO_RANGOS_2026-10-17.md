# Diccionario de Datos, Columnas y Rangos — BDAS Multiusuario

> **Fecha de consolidación:** 17 de Octubre de 2026  
> **Libro Maestro:** `H:\ResidenciaBD\BDAS_Multiusuario.xlsm` (MD5: `9E9685D7F5F8E33B8A139EE4929AE543`)

---

## 1. Catálogo Completo de Hojas del Libro (49 Hojas)

Excel identifica internamente cada hoja mediante dos propiedades esenciales:
*   **`Name`**: Nombre visible en la pestaña (puede cambiar si un usuario lo renombra).
*   **`CodeName`**: Identificador interno de VBA (inmutable desde la interfaz de Excel, ej. `Hoja2`). **Siempre debe priorizarse el uso de CodeNames en el código VBA**.

| CodeName | Nombre Visible (`Name`) | Visibilidad | Propósito Funcional |
| :--- | :--- | :--- | :--- |
| **`Hoja43`** | `INICIO` | `Hidden` | Hoja de bienvenida y transición previa al login. |
| **`Hoja2`** | `RESIDENCIA GIJÓN` | `Visible` | **Hoja operativa activa:** Solicitudes y reservas de Gijón. |
| **`Hoja1`** | `Calendario GIJÓN` | `Visible` | Vista de cuadrícula visual de ocupación por habitación de Gijón. |
| **`Hoja7`** | `RESUMEN GIJÓN` | `Visible` | Estadísticas agregadas, contadores y facturación de Gijón. |
| **`Hoja20`** | `RESIDENCIA SOTO` | `Visible` | **Hoja operativa activa:** Solicitudes y reservas de Soto del Barco. |
| **`Hoja22`** | `Calendario SOTO` | `Visible` | Vista de cuadrícula visual de ocupación de apartamentos en Soto. |
| **`Hoja9`** | `RESUMEN SOTO` | `Visible` | Estadísticas agregadas, contadores y facturación de Soto. |
| **`Hoja24`** | `RESIDENCIA OVIEDO` | `Visible` | **Hoja operativa activa:** Solicitudes y reservas de Oviedo. |
| **`Hoja23`** | `Calendario OVIEDO` | `Visible` | Vista de cuadrícula visual de ocupación por habitación de Oviedo. |
| **`Hoja17`** | `RESUMEN OVIEDO` | `Visible` | Estadísticas agregadas, contadores y facturación de Oviedo. |
| **`Hoja34`** | `TARIFAS UNIDAD` | `Hidden` | Matriz de tarifas de Gijón y Oviedo con control anti-solapes por fechas. |
| **`Hoja28`** | `Turnos-Precios` | `Hidden` | Tarifas y precios por turno para los apartamentos de Soto del Barco. |
| **`Hoja41`** | `Festivos` | `Hidden` | Calendario laboral y festivos oficiales aplicables a cálculo de estancias. |
| **`Hoja45`** | `CONFIG_PWD` | `VeryHidden` | Registro dinámico de contraseñas del sistema ofuscadas con ROT1. |
| **`Hoja42`** | `CONFIG` | `VeryHidden` | Parámetros del sistema y configuración de búsqueda BDAS. |
| **`Hoja6`** | `USUARIOS` | `VeryHidden` | Tabla interna de usuarios y permisos (heredada de Excel). |
| **`Hoja32`** | `LOG_GIJÓN` | `Hidden` | Auditoría de accesos y modificaciones en Gijón. |
| **`Hoja39`** | `LOG_SOTO` | `Hidden` | Auditoría de accesos y modificaciones en Soto. |
| **`Hoja40`** | `LOG_OVIEDO` | `Hidden` | Auditoría de accesos y modificaciones en Oviedo. |
| **`Hoja44`** | `SOLAPES_ACEPTADOS_GIJON` | `VeryHidden` | Registro histórico de solapes autorizados en Gijón. |
| **`Hoja53`** | `SOLAPES_ACEPTADOS_OVIEDO` | `VeryHidden` | Registro histórico de solapes autorizados en Oviedo. |
| **`Hoja37`** | `Habitaciones Bloqueadas GIJÓN` | `Hidden` | Registro de habitaciones fuera de servicio por reformas/mantenimiento. |
| **`Hoja35`** | `Apartamentos Bloqueados SOTO` | `VeryHidden` | Registro de apartamentos bloqueados en Soto. |
| **`Hoja36`** | `Habitaciones Bloqueadas OVIEDO` | `VeryHidden` | Registro de habitaciones bloqueadas en Oviedo. |
| **`Hoja18`** | `FACTURACIÓN GIJÓN` | `Hidden` | Resumen de facturas emitidas por mes en Gijón. |
| **`Hoja30`** | `FACTURACIÓN SOTO` | `Hidden` | Resumen de facturas emitidas por mes en Soto. |
| **`Hoja25`** | `FACTURACIÓN OVIEDO` | `Hidden` | Resumen de facturas emitidas por mes en Oviedo. |
| **`Hoja4`** | `FACTURA GIJÓN` | `Hidden` | Plantilla imprimible de factura PAX para Gijón. |
| **`Hoja29`** | `FACTURA SOTO` | `Hidden` | Plantilla imprimible de factura PAX para Soto. |
| **`Hoja31`** | `FACTURA OVIEDO` | `Hidden` | Plantilla imprimible de factura PAX para Oviedo. |
| **`Hoja11`** | `SOBRE GIJÓN` | `VeryHidden` | Plantilla imprimible para ventanilla de sobre en Gijón. |
| **`Hoja54`** | `SOBRE OVIEDO` | `VeryHidden` | Plantilla imprimible para ventanilla de sobre en Oviedo. |
| **`Hoja12`** | `PLANTILLAS GIJÓN` | `VeryHidden` | Documentos tipo de concesión y resoluciones de Gijón. |
| **`Hoja33`** | `PLANTILLAS SOTO` | `VeryHidden` | Documentos tipo de concesión y resoluciones de Soto. |
| **`Hoja16`** | `PLANTILLAS OVIEDO` | `VeryHidden` | Documentos tipo de concesión y resoluciones de Oviedo. |
| **`Hoja27`** | `CONTRASEÑAS GIJÓN` | `VeryHidden` | Claves internas de alojamientos y WiFi de Gijón. |
| **`Hoja38`** | `CONTRASEÑAS SOTO` | `VeryHidden` | Claves internas de alojamientos y WiFi de Soto. |
| **`Hoja49`** | `Lavandería Gijón` | `VeryHidden` | Registro activo de servicio de lavandería en Gijón. |
| **`Hoja50`** | `HISTORICO_LAVANDERIA_GIJON` | `VeryHidden` | Histórico consolidado de lavandería Gijón. |
| **`Hoja21`** | `Lavandería Soto` | `VeryHidden` | Registro activo de lavandería en Soto. |
| **`Hoja19`** | `HISTORICO_LAVANDERIA_SOTO` | `VeryHidden` | Histórico consolidado de lavandería Soto. |
| **`Hoja51`** | `Lavandería Oviedo` | `VeryHidden` | Registro activo de lavandería en Oviedo. |
| **`Hoja52`** | `HISTORICO_LAVANDERIA_OVIEDO` | `VeryHidden` | Histórico consolidado de lavandería Oviedo. |
| **`Hoja8`** | `RESIDENCIA ESTUDIANTES` | `VeryHidden` | Módulo específico de gestión de estudiantes universitarios. |
| **`Hoja5`** | `FACTURA R. ESTUDIANTES` | `VeryHidden` | Plantilla de facturación de residentes estudiantes. |
| **`Hoja10`** | `LIQUIDACION R. ESTUDIANTES` | `VeryHidden` | Liquidación agrupada de estudiantes. |
| *(Eliminada)* | *(Anterior `LISTA NEGRA` / `Hoja3`)* | *Suprimida* | Hoja eliminada en Hito 2026-10-18 (Opción C). La gestión opera 100% desacoplada en memoria contra Access (`Residencia_BE.accdb`) mediante `frmGestionListaNegra`. |
| **`Hoja15`** | `Evaluación` | `VeryHidden` | Parámetros de baremación de solicitudes. |
| **`Hoja14`** | `Auxiliar` | `VeryHidden` | Tablas de apoyo y listas desplegables secundarias. |

---

## 2. Catálogo de Columnas de las Hojas de Residencia Operativas

> [!WARNING]
> **Discrepancia Crítica de Columnas entre Sedes:**  
> • **`RESIDENCIA GIJÓN`** y **`RESIDENCIA SOTO`** tienen **30 columnas** (`A` a `AD`). El importe calculado se ubica en la **Columna `U`**.  
> • **`RESIDENCIA OVIEDO`** tiene **33 columnas** (`A` a `AG`). Dispone de una columna adicional previa para camas supletorias (`S`). Por tanto, en Oviedo el importe calculado se ubica en la **Columna `V`**.  
> *Cualquier macro genérica debe tener en cuenta este desplazamiento según la sede.*

### Estructura de `RESIDENCIA GIJÓN` (`Hoja2`) — 30 Columnas

| Col | Encabezado Fila 1 | Tipo / Origen | Detalle / Fórmula Maestra |
| :---: | :--- | :--- | :--- |
| **A** | `Nº ORDEN` | Entero (Long) | Clave unívoca autonumérica coordinada con Access. |
| **B** | `FECHA PETICION` | Fecha | Fecha de registro de la solicitud. |
| **C** | `NºM FACT.` | Entero (Long) | Número de factura asignado correlativo por sede y ejercicio. |
| **D** | `FINALIDAD` | Texto (Lista) | `Descanso`, `Oficial`, `Particular`, `Comisión`, etc. |
| **E** | `EMPLEO` | Texto | Empleo / Categoría profesional del solicitante. |
| **F** | `SITUACION` | Texto | `Activo`, `Reserva`, `Retirado`, etc. |
| **G** | `EVALUACIÓN` | Texto | Baremación resultante de la solicitud. |
| **H** | `COMISIÓN` | Texto (`SI`/`NO`) | Indica si aplica tarifa oficial de comisión de servicio. |
| **I** | `DNI` | Texto | Documento nacional de identidad. |
| **J** | `RANGO` | Texto | Rango militar o de empleo. |
| **K** | `Nombre` | Texto | Nombre y apellidos del huésped principal. |
| **L** | `ENTRADA` | Fecha | Fecha de ingreso en la residencia. |
| **M** | `SALIDA` | Fecha | Fecha de desocupación. |
| **N** | `DIAS USO` | Entero | Número de días de pernocta. |
| **O** | `PAX` | Entero | Número total de personas en la reserva. |
| **P** | `RESOLUCION` | Texto | `CONCEDIDA`, `DENEGADA`, `DESESTIMADA`, `RENUNCIA`, etc. |
| **Q** | `HAB. IND.` | Entero | Cantidad de habitaciones individuales solicitadas. |
| **R** | `HAB. DOBLE` | Entero | Cantidad de habitaciones dobles solicitadas. |
| **S** | `NºM HAB.` | Texto | Habitaciones físicas asignadas (ej. `1, 4`). |
| **T** | `DTO. FAM. NUM.` | Texto | Bonificación: `DTO 10%`, `DTO 20%` o vacío. |
| **U** | `IMPORTE` | **Fórmula** | **Fórmula Dinámica Matricial vinculada a `TARIFAS UNIDAD`:**<br>`=SI(P2="";"";SI((1+SI(SUMAPRODUCTO(('TARIFAS UNIDAD'!$A$3:$A$7<=L2)*('TARIFAS UNIDAD'!$B$3:$B$7>=L2)*'TARIFAS UNIDAD'!$C$3:$C$7)=0;0,1;SUMAPRODUCTO(('TARIFAS UNIDAD'!$A$3:$A$7<=L2)*('TARIFAS UNIDAD'!$B$3:$B$7>=L2)*'TARIFAS UNIDAD'!$C$3:$C$7)))*(SI(O(P2="CONCEDIDA";P2="SI";P2="REEVALUADA");1;0)*SI(T2="DTO 20%";0,8;1)*SI(T2="DTO 10%";0,9;1)*(SI(H2="NO";N2*(SI(ESNUMERO(Q2);Q2;0)*SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=L2)*('TARIFAS UNIDAD'!$F$3:$F$9>=L2)*'TARIFAS UNIDAD'!$G$3:$G$9)+SI(ESNUMERO(R2);R2;0)*SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=L2)*('TARIFAS UNIDAD'!$F$3:$F$9>=L2)*'TARIFAS UNIDAD'!$I$3:$I$9));0)+SI(H2="SI";N2*(SI(ESNUMERO(Q2);Q2;0)*SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=L2)*('TARIFAS UNIDAD'!$F$3:$F$9>=L2)*'TARIFAS UNIDAD'!$H$3:$H$9)+SI(ESNUMERO(R2);R2;0)*SUMAPRODUCTO(('TARIFAS UNIDAD'!$E$3:$E$9>0)*('TARIFAS UNIDAD'!$E$3:$E$9<=L2)*('TARIFAS UNIDAD'!$F$3:$F$9>=L2)*'TARIFAS UNIDAD'!$J$3:$J$9));0)))/SI($AD2="SI";MAX(1;$O2);1)=0;"";...))` |
| **V** | `TELEFONO` | Texto | Teléfono de contacto. |
| **W** | `DIRECCIÓN` | Texto | Domicilio habitual. |
| **X** | `CP` | Texto | Código postal (5 dígitos). |
| **Y** | `POBLACIÓN` | Texto | Municipio. |
| **Z** | `PROVINCIA` | Texto | Provincia de residencia. |
| **AA** | `PAGADO` | Texto | Estado de liquidación (`SI`, `NO`, etc.). |
| **AB** | `GRABACIÓN SOLICITUD` | **Fórmula** | **Fórmula Generadora de Nombre de Archivo PDF:**<br>`=SI(O(A2=""; I2=""; K2=""); ""; TEXTO(A2; "0000") & ".S." & I2 & "." & K2 & ".pdf")` |
| **AC** | `COPIAR TEXTO EMAIL ` | **Fórmula** | `=SI(Y(A2<>""; P2<>""); "Nº  " & TEXTO(A2;"0") & " - " & P2; "")` |
| **AD** | `REPARTIR COSTE  HAB.` | Texto (`SI`/`NO`) | Divide el coste total entre el número de huéspedes de la habitación. |

---

### Estructura de `RESIDENCIA SOTO` (`Hoja20`) — 30 Columnas

| Col | Encabezado Fila 1 | Tipo | Descripción |
| :---: | :--- | :---: | :--- |
| **A a P** | `Nº ORDEN` hasta `RESOLUCION` | Datos | Idénticos a Gijón (en Col H el campo almacena `TURNO`). |
| **Q** | `PRECIO APTO.` | **Fórmula** | **Fórmula de Precio Base por Turno/Comisión:** Busca el importe en `'Turnos-Precios'!$A$1:$F$22` o calcula comisión. |
| **R** | `SUPLE. OCUPAN.` | **Fórmula** | **Fórmula de Suplemento de Ocupantes:** Aplica suplemento para estancias cortas (<5 días) o PAX adicional. |
| **S** | `APARTAMENTO` | Texto | Apartamento asignado (mapeado de `HabitacionesAsignadas` en Access). |
| **T** | `DTO. FAM. NUM.` | Texto | Bonificación familiar. |
| **U** | `IMPORTE` | **Fórmula** | **Fórmula de Liquidación de Soto:** Calcula el coste final ponderando `Q`, `R`, días, IVA y descuento. |
| **V a AA**| `TELEFONO` a `PAGADO` | Datos | Datos de contacto y pago idénticos a Gijón. |
| **AB** | `GRABACIÓN SOLICITUD` | **Fórmula** | **Fórmula Generadora de Nombre PDF:** `=SI(O(A2=""; I2=""; K2=""); ""; TEXTO(A2; "000") & ".S." & I2 & "." & K2 & ".pdf")` |
| **AC** | `COPIAR TEXTO EMAIL ` | **Fórmula** | Texto formateado para correos (`Nº X - ESTADO`). |
| **AD** | `REPARTIR COSTE  APTO.`| Texto | Indicador de división de coste. |

---

### Estructura de `RESIDENCIA OVIEDO` (`Hoja24`) — 33 Columnas

| Col | Encabezado Fila 1 | Diferencia vs Gijón |
| :---: | :--- | :--- |
| **A a R** | `Nº ORDEN` hasta `HAB. DOBLE` | Idénticas a Gijón. |
| **S** | `CAMA SUPLE.` | **Exclusiva Oviedo:** Cama supletoria solicitada. |
| **T** | `NºM HAB.` | Desplazada (en Gijón es S). |
| **U** | `DTO. FAM. NUM.` | Desplazada (en Gijón es T). |
| **V** | `IMPORTE` | **Fórmula Maestra de Oviedo** calculada contra el tramo `TARIFAS UNIDAD!$E$13:$F$19` y ponderando camas supletorias (`S2` contra columna `K` o `L`). |
| **W a AB** | `TELEFONO` a `PAGADO` | Desplazadas una posición respecto a Gijón. |
| **AC** | `REPARTIR COSTE  HAB.` | Indicador de división de coste. |
| **AD** | `TAG` | Campo de etiquetado interno de Oviedo. |
| **AE** | `R` | Indicador de reserva rápida. |
| **AF** | `GRABACIÓN SOLICITUD` | **Fórmula Generadora de Nombre de Archivo PDF:** `=SI(O(A2=""; I2=""; K2=""); ""; TEXTO(A2; "0000") & ".S." & I2 & "." & K2 & ".pdf")` |
| **AG** | `COPIAR TEXTO EMAIL ` | **Fórmula:** Texto preparado para comunicaciones (`Nº X - ESTADO`). |

---

## 3. Matriz de Configuración: `TARIFAS UNIDAD` (`Hoja34`)

La hoja se divide en dos grandes bloques tabulares gobernados por el evento anti-solapes [Hoja34_TarifasUnidad.cls](file:///H:/ResidenciaApp/bdas-multiusuario/vba-modules/Hoja34_TarifasUnidad.cls):

```
      Col A      Col B      Col C        Col E   Col F      Col G             Col H          Col I         Col J          Col K             Col L               Col M                Col N             Col O
   ┌──────────┬──────────┬─────────┐   ┌───────┬───────┬────────────────┬────────────────┬───────────┬────────────────┬─────────────────┬───────────────────┬────────────────────┬────────────────────┬─────────┐
 1 │ IVA      │          │         │   │ TARIFAS GIJÓN (sin IVA)                                                                                                                                           │
 2 │ Desde    │ Hasta    │ Valor   │   │ Desde │ Hasta │ Hab.Individual │ Hab.Ind-COMIS. │ Hab.Doble │ Hab.Dob-COMIS. │ Cama supletoria │ Cama suple-COMIS. │ Res.Estud.-CLASE A │ Res.Estud.-CLASE B │ Nº BOGC │
 3 │ 01/01/24 │ 31/12/99 │ 0.10    │   │ 01/01 │ 31/12 │     18.00      │     16.00      │   20.00   │     18.00      │      6.00       │       4.00        │       120.00       │       100.00       │         │
   └──────────┴──────────┴─────────┘   └───────┴───────┴────────────────┴────────────────┴───────────┴────────────────┴─────────────────┴───────────────────┴────────────────────┴────────────────────┴─────────┘
                                       ┌─────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────────┐
11                                     │ TARIFAS OVIEDO (sin IVA)                                                                                                                                    │
12                                     │ Desde │ Hasta │ Hab.Individual │ Hab.Ind-COMIS. │ Hab.Doble │ Hab.Dob-COMIS. │ Cama supletoria │ Cama suple-COMIS. │ Res.Estud.-CLASE A │ Res.Estud.-CLASE B │ Nº BOGC │
13                                     │ 01/01 │ 31/12 │     18.00      │     16.00      │   20.00   │     18.00      │      6.00       │       4.00        │       120.00       │       100.00       │         │
                                       └───────┴───────┴────────────────┴────────────────┴───────────┴────────────────┴─────────────────┴───────────────────┴────────────────────┴────────────────────┴─────────┘
```

*   **Tabla IVA:** Rango `$A$2:$C$7` (evalúa si la fecha de entrada cae en el periodo de IVA configurado).
*   **Tabla Gijón:** Rango `$E$3:$O$9` (hasta 7 tramos temporales con validación anti-solape).
*   **Tabla Oviedo:** Rango `$E$13:$O$19` (hasta 7 tramos temporales con validación anti-solape).

---

## 4. Catálogo de Nombres de Rango Globales (`Defined Names`)

El libro dispone de 90 rangos nombrados en Excel, destacando los siguientes grupos operativos:

1.  **Rangos de Impresión (`Print_Area`):**
    *   `'FACTURA GIJÓN'!$A$1:$V$42`
    *   `'FACTURA OVIEDO'!$A$1:$V$42`
    *   `'FACTURA SOTO'!$A$1:$V$42`
    *   `'FACTURA R. ESTUDIANTES'!$A$2:$U$39`
    *   `'LIQUIDACION R. ESTUDIANTES'!$A$1:$F$24`
    *   `'SOBRE GIJÓN'!$A$1:$G$24`
    *   `'SOBRE OVIEDO'!$A$1:$F$16`
2.  **Disponibilidad de Alojamientos:**
    *   `HabitacionesDisponiblesGIJON`: `"1, 2, 3, 4, 5, 6, 7, Of.1, Of.2, Of.3, Est.1, Est.2, Est.3"`
    *   `HabitacionesDisponiblesOVIEDO`: `"4, 7, 8, 9, 13, 14"`
    *   `ApartamentosDisponibles`: `"Ap.2, Ap.4, Ap.5, Ap.6, Ap.7, Ap.8, Ap.9, Ap.10, Ap.12"`
3.  **Rangos Dinámicos de Consulta y Facturación:**
    *   `FACTURACION_PLAZA`: Rangos de columnas de facturación por plaza en Gijón (`$E:$AD`), Soto (`$E:$AD`) y Oviedo (`$E:$AH`).
    *   `FestivosLista`: `Festivos!$A$20` (lista maestra consumida por macros de calendario).
    *   `CodigosParaPoblacion` y `PoblacionesParaCodigo`: Fórmulas dinámicas con `DESREF` y `COINCIDIR` para el autocompletado en cascada de códigos postales y municipios.
