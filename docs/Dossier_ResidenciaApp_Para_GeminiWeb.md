# DOSSIER TÉCNICO Y FUNCIONAL DEL SISTEMA RESIDENCIAAPP
## Documento Maestro de Contexto, Arquitectura, Flujos y Catálogo de Mejoras para Inteligencia Artificial

> **CLASIFICACIÓN DEL DOCUMENTO:** Documento de Contexto Técnico y Funcional 100% Anonimizado.  
> **PRIVACIDAD Y RGPD:** Todos los nombres personales, DNI, teléfonos, correos y domicilios incluidos en este documento son ficticios, sintéticos o descriptivos. No contiene datos personales reales.  
> **FECHA DE CONSOLIDACIÓN:** 22 de octubre de 2026.  
> **VERSION DE REFERENCIA BASE:** Hito 2026-10-21 23:30 (Versión Oficial Vigente).

---

## 0. PROMPT DIRECTIVO DE INICIO PARA GEMINI WEB (INSTRUCCIÓN PARA EL MODELO)

```
[ROL DEL ASISTENTE]:
Actúa como un Arquitecto de Software Senior, Especialista en Ingeniería de Sistemas Distribuidos,
Consultor de UX/UI en entornos corporativos y Experto en Microsoft Excel VBA y Bases de Datos Relacionales (Access / SQL Server).

[CONTEXTO QUE SE TE PROPORCIONA]:
Este documento contiene la especificación completa, arquitectura de persistencia, mapa de hojas,
esquema de base de datos relacional, flujos operativos paso a paso, código fuente de referencia
y catálogo de propuestas de mejora de "ResidenciaApp" (Sistema de Gestión Integral de Residencias
de Descanso de la Guardia Civil en Asturias: sedes de Gijón, Soto del Barco y Oviedo).

[TUS OBJETIVOS DE ANÁLISIS]:
1. Conocer a la perfección todos los flujos de la aplicación (reservas, asignación de habitaciones,
   calendarios, facturación PAX y estudiantes, control de cobros, lavandería y auditoría).
2. Debatir y proponer optimizaciones de alto nivel sobre los retos actuales del sistema:
   - Proyecto SICO-N2: Continuidad operativa fuera de línea mediante búfer local no volátil y sincronización FIFO.
   - Ergonomía visual: Mitigación de parpadeos gráficos y modernización de interfaces sin romper estabilidad.
   - Escalabilidad y transición gradual hacia la versión Web (React + .NET 8).
```

---

## 1. VISIÓN GENERAL DEL SISTEMA Y DOMINIO DE NEGOCIO

### 1.1. Objeto y Alcance de la Aplicación
**ResidenciaApp (BDAS Multiusuario)** es el sistema informático corporativo utilizado por la Guardia Civil para gestionar el ciclo de vida completo de ocupación, cobro, facturación y auditoría de sus residencias y casas cuartel de descanso en el Principado de Asturias.

El sistema administra tres residencias principales, cada una con particularidades operativas propias:
1. **Sede GIJÓN (Casa Cuartel de Gijón):** Régimen general de descanso, habitaciones individuales y dobles, plazas estables para estudiantes/cursos durante el año académico, y servicio de comidas/PAX.
2. **Sede SOTO DEL BARCO:** Régimen estival especializado por **turnos cerrados de 10 días** (Turno 1, Turno 2, Turno 3, etc.) y apartamentos/habitaciones vacacionales.
3. **Sede OVIEDO:** Régimen general de alojamiento, plazas de estudiantes y soporte para camas supletorias por habitación.

### 1.2. Tipología de Huéspedes y Estancias
- **Régimen General de Descanso:** Personal del cuerpo (oficiales, suboficiales, cabos y guardias, retirados, viudos/as) y sus beneficiarios.
- **Comisión de Servicio:** Personal desplazado por motivos de servicio oficial.
- **Estudiantes / Opositores:** Plazas mensuales de curso lectivo con tarifa fija y facturación específica.
- **Transeúntes:** Pernoctas de corta duración.
- **Acompañantes (PAX):** Conteo de ocupantes reales para cálculo de tarifas, desayunos y partes de estancia.

### 1.3. Reglas de Negocio Clave
1. **Unicidad de la Orden:** Cada reserva tiene un **Número de Orden correlativo y único**. En entorno multiusuario, este número se genera de manera atómica en la base de datos central Access mediante contador autonumérico para evitar colisiones.
2. **Control de Duplicidades (Ventana de 30 días):** El sistema impide o alerta si un mismo DNI solicita alojamiento en fechas idénticas o solapadas dentro de una ventana temporal reciente en cualquiera de las residencias.
3. **Control de Lista Negra (Vetos):** Comprobación en tiempo real contra la tabla de personas no autorizadas o con incidencias graves previas.
4. **Tarifas Dinámicas vinculadas a Fechas:** Los precios del alojamiento no están fijos en código; se calculan dinámicamente evaluando la fecha de estancia contra los tramos temporales definidos en la hoja maestra `TARIFAS UNIDAD`.
5. **Separación de Sedes y Roles:** Un operador solo puede acceder y operar las residencias que tiene expresamente asignadas en su perfil de usuario (`GIJON`, `SOTO`, `OVIEDO` o combinaciones). Solo el rol `ADMIN` tiene acceso transversal y permisos para gestionar usuarios y RGPD.

---

## 2. ARQUITECTURA TÉCNICA CLIENTE-SERVIDOR

El sistema está estructurado bajo una **Arquitectura Cliente-Servidor Transaccional en Tiempo Real sobre LAN/Unidad Compartida**:

```
+-----------------------------------------------------------------------------------+
|                        CLIENTES FRONT-END (Múltiples PCs)                         |
|                                                                                   |
|  +---------------------------+             +---------------------------+          |
|  | Puesto Operador 1 (Gijón) |             | Puesto Operador 2 (Oviedo)|          |
|  | BDAS_Multiusuario.xlsm    |             | BDAS_Multiusuario.xlsm    |          |
|  +---------------------------+             +---------------------------+          |
|                \                                         /                        |
|                 \                                       /                         |
|                  v                                     v                          |
|             Conexiones OLEDB ADODB (Short-lived: Open -> Exec -> Close)           |
+-----------------------------------------------------------------------------------+
                                         |
                                         v
+-----------------------------------------------------------------------------------+
|                     SERVIDOR CENTRAL DE ARCHIVOS (Unidad H:\)                     |
|                                                                                   |
|   Ubicación: H:\ResidenciaBD\Residencia_BE.accdb                                |
|   - Motor Relacional: Microsoft Access (ACE.OLEDB.12.0)                           |
|   - Tablas Maestras: Ordenes, Usuarios, LogActividad, ListaNegra, Tarifas        |
|   - Bloqueo a nivel de registro (Pessimistic / Optimistic Record Locking)         |
|   - Copia de Despliegue en Red: H:\ResidenciaBD\BDAS_Multiusuario.xlsm           |
|   - Repositorio Central de Plantillas y Versiones de Referencia protegidas       |
+-----------------------------------------------------------------------------------+
```

### 2.1. Front-End: Microsoft Excel VBA (`BDAS_Multiusuario.xlsm`)
- **Libro binario macro-habilitado (.xlsm):** Contiene 50 hojas de cálculo estructuradas y 171 componentes VBA (módulos estándar `.bas`, módulos de clase de hoja `.cls`, y formularios `.frm`).
- **Cinta de Opciones Personalizada (Ribbon XML):** Define dos pestañas completas en la interfaz superior de Office:
  * `GESTIÓN RESIDENCIAS ASTURIAS`: Alojamientos, Calendarios, Facturación, Lavandería y Utilidades.
  * `AUDITORÍA Y SEGURIDAD`: Hojas LOG por sede, Selector de Residencia, Panel RGPD y Salida Segura.
  * Oculta las pestañas nativas de Excel para usuarios con rol `OPERADOR`.
- **Persistencia de Sesión Inmune:** El usuario autenticado y su rol se almacenan de forma persistente y auto-reparable en Nombres Definidos Ocultos de Excel (`app_session_user`), haciéndola inmune a reseteos de variables de memoria de VBA.

### 2.2. Back-End: Base de Datos Relacional (`Residencia_BE.accdb`)
- Archivo central ubicado en `H:\ResidenciaBD\Residencia_BE.accdb`.
- Patrón de Conexión de Vida Corta (*Short-Lived Connection*): La aplicación nunca mantiene conexiones ADO permanentemente abiertas hacia el archivo `.accdb` para evitar corrupciones de archivo `.ldb` ante cierres inesperados de red. Abre la conexión, ejecuta la consulta o transacción (INSERT/UPDATE/SELECT) y cierra inmediatamente.

### 2.3. Gobernanza Técnica en 5 Capas (Estándar ECRMC)
1. **Regla Estricta del Literal `GIJÓN`:** Prohibido el uso de secuencias desfasadas tipo `Chr(211)`.
2. **Protocolo Cero Mojibake:** Codificación nativa Windows-1252 ANSI obligatoria con escáner pre-commit en Git.
3. **Protección contra Inyección de Fórmulas:** Los valores de auditoría o entradas que inician con el carácter igual (`=`) se prefijan automáticamente con apóstrofo (`'=`) neutralizando el Error 1004.
4. **Borrado Físico Exclusivo RGPD:** Ninguna acción ordinaria (como borrar una fila o pulsar Supr) borra registros en Access; solo archiva o desvincula la vista. La destrucción física de datos queda reservada a la pestaña de RGPD.
5. **Inmutabilidad del Golden Master:** La réplica maestra original permanece en solo lectura como póliza de seguro de arranque.

---

## 3. MAPA COMPLETO DE HOJAS DE EXCEL (LAS 50 HOJAS DEL SISTEMA)

| Nº Hoja | CodeName | Nombre Visible | Estado Inicial | Propósito y Funcionalidad |
|:---:|:---:|:---|:---:|:---|
| **1** | `Hoja43` | `INICIO` | Visible al arrancar | Carátula histórica de la Casa Cuartel con auto-fit dinámico y tarjeta flotante de Login (`frmLogin`). |
| **2** | `Hoja2` | `RESIDENCIA GIJÓN` | Visible (si autorizado) | Matriz principal de solicitudes y reservas de Gijón (Columnas A a AB). |
| **3** | `Hoja44` | `SOLAPES_ACEPTADOS_GIJON`| VeryHidden | Registro interno de solapamientos autorizados con justificación. |
| **4** | `Hoja1` | `Calendario GIJÓN` | Visible | Cuadrante interactivo de ocupación mensual/diaria por habitaciones de Gijón. |
| **5** | `Hoja7` | `RESUMEN GIJÓN` | Visible | Resumen estadístico mensual de ocupación, pagos e importes de Gijón. |
| **6** | `Hoja45` | `CONFIG_PWD` | VeryHidden | Almacenamiento cifrado de contraseñas de protección interna. |
| **7** | `Hoja42` | `CONFIG` | VeryHidden | Parámetros generales y rutas del sistema. |
| **8** | `Hoja6` | `USUARIOS` | VeryHidden | Espejo local de permisos y usuarios sincronizados desde Access. |
| **9** | `Hoja10` | `FACTURA GIJÓN` | Visible | Plantilla de emisión de factura oficial PAX vinculada a tarifas de Gijón. |
| **10** | `Hoja11` | `SOBRE GIJÓN` | Visible | Impresión de sobre de correspondencia oficial con franqueo y remitente. |
| **11** | `Hoja12` | `FACTURA ESTUDIANTE GIJÓN`| Visible | Factura mensual para estudiantes de curso académico en Gijón. |
| **12** | `Hoja13` | `LAVANDERÍA GIJÓN` | Visible | Control diario de prendas y sábanas por habitación en Gijón. |
| **13** | `Hoja46` | `LOG_GIJÓN` | Visible (Ribbon) | Auditoría local estructurada en 6 columnas de todos los cambios en Gijón. |
| **14** | `Hoja49` | `HISTORICO_LAVANDERIA_GIJON`| VeryHidden | Archivo acumulado de partes de lavandería de Gijón. |
| **15** | `Hoja3` | `RESIDENCIA SOTO` | Visible (si autorizado) | Matriz principal de reservas de Soto del Barco (gestión de Turnos estivales). |
| **16** | `Hoja8` | `Calendario SOTO` | Visible | Cuadrante visual de ocupación de apartamentos y habitaciones en Soto. |
| **17** | `Hoja14` | `RESUMEN SOTO` | Visible | Resumen estadístico mensual de ocupación y cobros de Soto. |
| **18** | `Hoja17` | `FACTURA SOTO` | Visible | Plantilla de facturación oficial para estancias en Soto del Barco. |
| **19** | `Hoja20` | `LAVANDERÍA SOTO` | Visible | Control de lencería y mudas de Soto del Barco. |
| **20** | `Hoja47` | `LOG_SOTO` | Visible (Ribbon) | Auditoría local estructurada de todos los movimientos de Soto. |
| **21** | `Hoja50` | `HISTORICO_LAVANDERIA_SOTO`| VeryHidden | Archivo histórico de lavandería de Soto. |
| **22** | `Hoja4` | `RESIDENCIA OVIEDO` | Visible (si autorizado) | Matriz principal de reservas de Oviedo (incluye camas supletorias). |
| **23** | `Hoja9` | `Calendario OVIEDO` | Visible | Cuadrante visual de ocupación de habitaciones de Oviedo. |
| **24** | `Hoja21` | `RESUMEN OVIEDO` | Visible | Resumen estadístico mensual de Oviedo. |
| **25** | `Hoja24` | `FACTURA OVIEDO` | Visible | Plantilla de facturación oficial de Oviedo. |
| **26** | `Hoja25` | `SOBRE OVIEDO` | Visible | Impresión de sobre de correspondencia oficial de Oviedo. |
| **27** | `Hoja26` | `FACTURA ESTUDIANTE OVIEDO`| Visible | Factura mensual de estudiantes para la sede de Oviedo. |
| **28** | `Hoja27` | `LAVANDERÍA OVIEDO` | Visible | Control de lencería y mudas de Oviedo. |
| **29** | `Hoja48` | `LOG_OVIEDO` | Visible (Ribbon) | Auditoría local estructurada de todos los movimientos de Oviedo. |
| **30** | `Hoja51` | `HISTORICO_LAVANDERIA_OVIEDO`| VeryHidden | Archivo histórico de lavandería de Oviedo. |
| **31** | `Hoja34` | `TARIFAS UNIDAD` | Visible (Admin) | Tabla maestra de tarifas oficiales por tramos temporales con fórmulas matriciales. |
| **32** | `Hoja35` | `FESTIVOS` | VeryHidden | Listado de días festivos oficiales diferenciados por municipio/sede. |
| **33** | `Hoja36` | `HABITACIONES` | VeryHidden | Catálogo físico de habitaciones por sede (individuales, dobles, camas). |
| **34** | `Hoja37` | `LISTAS_DESPLEGABLES` | VeryHidden | Listas maestras de validación de celdas (Empleos, Estados, Resoluciones). |
| **35** | `Hoja38` | `CP_POBLACIONES` | VeryHidden | Base de datos de Códigos Postales, Municipios y Provincias de España. |
| **36** | `Hoja39` | `LISTA_NEGRA` | VeryHidden | Espejo local de personas con restricción de acceso. |
| **37** | `Hoja40` | `RGPD_PANEL` | Visible (Admin) | Panel de control de borrado y anonimización según RGPD. |
| **38-50**| Varias | Plantillas auxiliares | VeryHidden | Plantillas de importación, respaldo y buffers de cálculo intermedio. |

---

## 4. DICCIONARIO DE DATOS RELACIONAL ACCESS (`Residencia_BE.accdb`)

### 4.1. Tabla Maestra: `Ordenes` (41 Campos)
Almacena todas las adjudicaciones y reservas vivas de las tres residencias.

| Campo | Tipo Access | Tamaño | Descripción y Reglas de Negocio |
|:---|:---:|:---:|:---|
| `NumOrden` | Entero Largo | 4 | **Clave de Orden**. Autonumérico en Access, asignado atómicamente. |
| `Residencia` | Texto | 20 | Sede de la reserva: `GIJON`, `SOTO` u `OVIEDO`. |
| `FechaPeticion` | Fecha/Hora | 8 | Fecha en la que el usuario registró o solicitó la estancia. |
| `NumFactura` | Entero Largo | 4 | Número oficial de factura emitido. Null si no está facturada. |
| `Finalidad` | Texto | 50 | Motivo: `DESCANSO`, `COMISIÓN DE SERVICIO`, `ESTUDIANTE`, `TRANSEÚNTE`. |
| `Empleo` | Texto | 50 | Grado militar: `GUARDIA CIVIL`, `CABO`, `SARGENTO`, `TENIENTE`, etc. |
| `Situacion` | Texto | 50 | Situación administrativa: `ACTIVO`, `RESERVA`, `RETIRADO`, `VIUDEDAD`. |
| `Evaluacion` | Texto | 50 | Baremos de concesión: `CONCEDIDA`, `DENEGADA`, `EN ESPERA`. |
| `Turno` | Texto | 50 | Turno estival (específico de Soto del Barco: `TURNO 1`, `SIN TURNO`, etc.). |
| `Comision` | Texto | 50 | Indicador de si es comisión de servicio indemnizable (`SÍ` / `NO`). |
| `DNI` | Texto | 20 | Documento de identidad del titular (validado con letra algorítmica). |
| `Rango` | Texto | 50 | Escala o categoría del titular. |
| `Nombre` | Texto | 100 | Nombre de pila del titular. |
| `Apellidos` | Texto | 100 | Apellidos completos del titular. |
| `FechaEntrada` | Fecha/Hora | 8 | Fecha de inicio del alojamiento (Check-in). |
| `FechaSalida` | Fecha/Hora | 8 | Fecha de fin del alojamiento (Check-out). |
| `DiasUso` | Entero | 2 | Días netos pernoctados (`FechaSalida - FechaEntrada`). |
| `PAX` | Entero | 2 | Número total de ocupantes reales (titular + acompañantes). |
| `Resolucion` | Texto | 30 | `CONCEDIDA`, `DENEGADA`, `RENUNCIA`, `ANULADA`. |
| `NumHabIndividuales`| Entero | 2 | Cantidad de habitaciones individuales requeridas. |
| `NumHabDobles` | Entero | 2 | Cantidad de habitaciones dobles requeridas. |
| `CamaSuple` | Texto | 50 | Uso de cama supletoria (`SÍ` / `NO` o número de camas). |
| `HabitacionesAsignadas`| Texto | 200 | Números de habitación asignados (ej: `101`, `204-205`). |
| `DtoFamNum` | Texto | 50 | Descuento familia numerosa: `NO`, `GENERAL (20%)`, `ESPECIAL (50%)`. |
| `Importe` | Doble | 8 | Importe total liquidado en euros (cuadre al céntimo). |
| `Telefono` | Texto | 20 | Teléfono de contacto directo del huésped. |
| `Direccion` | Texto | 200 | Dirección postal del domicilio del huésped. |
| `CodigoPostal` | Texto | 10 | Código postal de 5 dígitos (autocompleta Municipio y Provincia). |
| `Poblacion` | Texto | 100 | Municipio o localidad de residencia. |
| `Provincia` | Texto | 50 | Provincia asociada al código postal. |
| `EstadoPago` | Texto | 30 | Estado contable: `METÁLICO`, `BANCO`, `TARJETA`, `BIZUM`, `PENDIENTE`. |
| `PagadoResidencia` | Texto | 30 | Indicador de cobro en la propia sede. |
| `Observaciones` | Memo/Texto Largo| - | Notas internas, incidencias o peticiones de cuna/cama. |
| `FechaCreacion` | Fecha/Hora | 8 | Marca de tiempo exacta del sistema al crearse el registro. |
| `UsuarioCreacion` | Texto | 50 | Nombre del operador que grabó la solicitud. |
| `Solapamiento` | Sí/No (Bool) | 1 | Flag de detección de solape con otra solicitud en las mismas fechas. |
| `ConsentimientoRGPD`| Sí/No (Bool) | 1 | Indicador de consentimiento de tratamiento de datos personales. |

### 4.2. Tabla: `Usuarios` (Control de Acceso y Roles)
- `Id` (Autonumérico, Clave Primaria).
- `NombreUsuario` (Texto 50, Único): Login del operador (ej: `operador_gijon`, `admin`).
- `Clave` (Texto 128): Hash criptográfico seguro **SHA-256** de la contraseña.
- `NombreCompleto` (Texto 100): Nombre legible para encabezados y barras de estado.
- `Rol` (Texto 20): `ADMIN` (acceso a todo) u `OPERADOR` (acceso restringido a sus sedes).
- `Residencias` (Texto 60): Lista separada por comas de sedes autorizadas (ej: `GIJON,SOTO`).
- `Activo` (Sí/No): Permite revocar el acceso instantáneamente.
- `FechaAlta` (Fecha/Hora): Auditoría de creación del usuario.

### 4.3. Tabla: `LogActividad` (Auditoría Central Inmutable)
- `Id` (Autonumérico).
- `FechaHora` (Fecha/Hora): Marca de tiempo del evento.
- `Usuario` (Texto 50): Operador responsable.
- `Accion` (Texto 200): Tipo de evento (`LOGIN`, `MODIFICACION_ORDEN`, `ALTA_LISTA_NEGRA`, etc.).
- `Detalle` (Texto Largo): Desglose forense de campos modificados, valor previo y valor nuevo.

### 4.4. Tabla: `ListaNegra` (Vetos y Restricciones)
- `Id` (Autonumérico).
- `DNI` (Texto 20): Clave de búsqueda del veto.
- `Nombre` (Texto 150): Nombre de la persona restringida.
- `Motivo` (Texto Largo): Causa formal del veto o resolución de expulsión.
- `FechaAlta` (Fecha/Hora): Fecha del registro.
- `Activo` (Sí/No): Estado de vigencia del veto.

### 4.5. Tabla: `ContadorFacturas` (Garantía de Correlatividad Fiscal)
- `Id` (Autonumérico).
- `Residencia` (Texto 20): Sede de la serie (`GIJON`, `SOTO`, `OVIEDO`).
- `Ejercicio` (Entero): Año contable (ej: `2026`).
- `UltimoNumero` (Entero Largo): Último número oficial correlativo emitido en ese ejercicio.

---

## 5. FLUJOS OPERATIVOS PASO A PASO (END-TO-END)

```
[Arranque de Excel] -> [Test Conexión Access] -> [Auto-Fit Carátula] -> [frmLogin Modal]
         |
         v
[Selección de Sede] -> [Descarga Datos Top 500 Access] -> [Activación Hoja Operativa]
         |
         +---> [Grabación Solicitud] -> [Test Duplicados 30d] -> [UPSERT Access en vivo]
         |
         +---> [Asignar Habitación]  -> [Pintado Calendario]  -> [Validar Solapes]
         |
         +---> [Facturación PAX]     -> [Tarifas Dinámicas]   -> [Nº Factura Central]
         |
         +---> [Cobro en Recepción]  -> [Metálico / Banco]    -> [Sync Resumen y LOG]
         |
         v
[Cierre con Aspa X] -> [MsgBox Confirmación] -> [ScreenUpdating=False] -> [Salida Suave]
```

### Flujo 1: Arranque, Autenticación y Selector de Residencia
1. El usuario abre `BDAS_Multiusuario.xlsm`.
2. Se dispara el evento `Workbook_Open` en `ThisWorkbook.cls`.
3. **Paso 1 (Comprobación de enlace):** Ejecuta `modDatabase.ComprobarConexion()`. Si `H:\ResidenciaBD\Residencia_BE.accdb` no responde, muestra mensaje explicativo y cierra el libro sin guardar para proteger la integridad.
4. **Paso 2 (Carátula proporcional):** Pone visible la hoja `INICIO`, detecta `UsableWidth` y `UsableHeight` de la pantalla y reescala automáticamente la imagen histórica de la Casa Cuartel para que el tercio inferior sea 100% visible sin recortes en cualquier monitor (1080p, 4K, 125% escalado).
5. **Paso 3 (Tarjeta de Login):** Muestra el formulario modal `frmLogin`, anclado milimétricamente en el patio empedrado entre los dos guardias civiles.
6. **Paso 4 (Validación SHA-256):** El operador introduce usuario y clave. Se calcula el hash SHA-256 y se compara contra la tabla `Usuarios` en Access.
7. **Paso 5 (Selector de Residencia):** Si el usuario tiene asignada una sola sede, se activa directamente. Si tiene varias (o es ADMIN), aparece `frmSelectorResidencia` para elegir sede activa.
8. **Paso 6 (Sincronización inicial):** Se descargan en memoria los 500 registros más recientes de Access y se activan únicamente las hojas autorizadas para esa sede.

### Flujo 2: Grabación de Solicitudes y Prevención de Duplicidades
1. El operador se posiciona en la última fila vacía de la hoja operativa (ej: `RESIDENCIA GIJÓN`).
2. Introduce DNI y fechas.
3. Se activa el motor `EvitarDuplicidadSolicitudesGyS`:
   - Consulta Access buscando si ese mismo DNI ha tenido reservas en los últimos 30 días o si coincide en fechas con otra solicitud.
   - Si detecta duplicado, muestra un diálogo interactivo con 3 opciones:
     * `Opción 1 -> Mantener ambas solicitudes`.
     * `Opción 2 -> Reemplazar solicitud previa`.
     * `Opción 3 -> Cancelar grabación`.
4. Comprueba la tabla `ListaNegra`. Si el DNI está vetado, bloquea inmediatamente la adjudicación.
5. Asigna el siguiente **Nº de Orden atómico** en Access e inserta la fila mediante transacción ADODB.
6. Mediante la API de Windows se normalizan nombres y poblaciones (conversión automática ProperCase, cálculo de letra de DNI y autocompletado de Código Postal).

### Flujo 3: Asignación de Habitaciones y Calendarios de Ocupación
1. Una vez aprobada la solicitud, el operador asigna número de habitación en la columna correspondiente (columna AB en Gijón).
2. Se valida que la habitación exista en el catálogo de la sede (`HABITACIONES`) y que su tipología (individual, doble, supletoria) sea coherente con los PAX.
3. El módulo de calendario (`ModuloCalendarioGijon.bas`, `Soto`, `Oviedo`):
   - Actualiza la cuadrícula mensual de ocupación.
   - Colorea las celdas según estado: Ocupado (verde/azul según tipo), Bloqueado (rojo), Festivo (sombreado especial).
   - Detecta solapes visuales si dos reservas intentan pisar la misma cama.
4. Botones de navegación temporal del Ribbon permiten avanzar o retroceder 5 o 10 días o saltar directamente a la fecha actual (`Date`).

### Flujo 4: Facturación PAX y Liquidación con Tarifas Dinámicas
1. Al completarse la estancia, el operador pulsa el botón **`Factura PAX`** en el Ribbon.
2. El sistema consulta en Access el contador oficial de la sede (`ContadorFacturas`) y reserva de forma atómica el siguiente número correlativo oficial (ej: `2026/0142`).
3. Se copian los datos a la hoja `FACTURA GIJÓN`.
4. **Cálculo de Tarifas Dinámicas:**
   - En lugar de aplicar precios estáticos, la hoja evalúa la fecha de estancia (`F19`) contra los tramos temporales de `TARIFAS UNIDAD` usando fórmulas matriciales `SUMAPRODUCTO`.
   - Aplica los coeficientes de reducción por Familia Numerosa (General 20%, Especial 50%).
   - Añade cargos por cama supletoria o desayunos si procede.
   - Genera el cuadre contable exacto al céntimo.
5. El operador puede pulsar **`Imprimir Sobre`**, lo que transfiere los datos del huésped a la plantilla `SOBRE GIJÓN` adaptada a las medidas de ventanilla oficial de Correos.

### Flujo 5: Cobro en Recepción y Estados de Pago
1. Al momento del pago, el operador despliega la columna `EstadoPago` en la hoja de residencia y selecciona: `METÁLICO`, `BANCO`, `TARJETA` o `BIZUM`.
2. El evento `Worksheet_Change` intercepta la modificación y ejecuta `modDatabase.ActualizarOrden` enviando un `UPDATE Ordenes SET EstadoPago = 'METÁLICO' WHERE NumOrden = X`.
3. Se sincroniza bidireccionalmente con la hoja `RESUMEN`, actualizando los totales de caja del día y el desglose contable mensual.

### Flujo 6: Auditoría Forense y Hojas LOG
1. Cada cambio manual sobre celdas críticas (habitaciones, fechas, importes, estados) es interceptado por `Worksheet_Change`.
2. Se genera un apunte automático en la hoja LOG de la sede (`LOG_GIJÓN`, `LOG_SOTO`, `LOG_OVIEDO`) con 6 columnas normalizadas:
   `[Usuario | Fecha/Hora | Nº Orden | Columna Modificada | Valor Anterior | Valor Nuevo]`
3. Se replica el evento en la tabla central `LogActividad` de Access.
4. Las hojas LOG están protegidas contra escritura; solo el sistema puede escribir mediante desprotección/reprotección por código con contraseña segura.
5. Se aplica la directiva contra inyección de fórmulas (`'=`): si un usuario escribe un texto que empieza por `=`, se le antepone un apóstrofo para impedir que Excel lo interprete como fórmula dañina.

### Flujo 7: Gestión RGPD y Purgado Histórico
1. Las órdenes con más de 1 año de antigüedad son transferidas periódicamente a la tabla `Ordenes_Historico`.
2. En cumplimiento estricto del RGPD, la destrucción física de datos personales (nombres, DNI, teléfonos) está prohibida en las hojas de trabajo ordinarias.
3. Se accede exclusivamente a través de la pestaña de Ribbon `RGPD / AUDITORÍA` por parte del usuario `ADMIN`:
   - El formulario `frmPanelRGPD` permite seleccionar el ejercicio a purgar.
   - Ejecuta consultas de actualización en Access sustituyendo datos personales por marcas de anonimización (`ANONIMIZADO_RGPD`), conservando únicamente los datos estadísticos y económicos disociados para cuadres de intervención militar.

### Flujo 8: Cierre Suave y Salida Segura
1. Al hacer clic en el aspa superior (X) de Excel o pulsar "Salir":
2. El evento `Workbook_BeforeClose` en `ThisWorkbook.cls` intercepta el evento.
3. Presenta un cuadro interactivo: `¿Desea salir de la aplicación?` con botón predeterminado en "No" para evitar salidas accidentales.
4. Si el usuario confirma "Sí":
   - Se desactiva el parpadeo de pantalla (`Application.ScreenUpdating = False`).
   - Se minimiza suavemente la ventana.
   - Se guarda el libro de forma ordenada (`ThisWorkbook.Save`).
   - Si no hay otros libros abiertos en la sesión de Excel, se cierra la instancia completa de Office sin dejar procesos zombi en segundo plano.

---

## 6. CÓDIGO FUENTE ESENCIAL Y PATRONES DE ARQUITECTURA

A continuación se incluyen fragmentos representativos y limpios de los componentes clave:

### 6.1. Patrón de Conexión de Vida Corta y Ejecución SQL (`modDatabase.bas`)
```vba
Private Const DB_PATH As String = "H:\ResidenciaBD\Residencia_BE.accdb"
Private Const PROVIDER_STRING As String = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="

Public Function GetConnectionString() As String
    GetConnectionString = PROVIDER_STRING & DB_PATH & ";Persist Security Info=False;"
End Function

Public Function ComprobarConexion() As Boolean
    Dim cn As Object
    On Error GoTo ErrorHandler
    Set cn = CreateObject("ADODB.Connection")
    cn.ConnectionTimeout = 5
    cn.Open GetConnectionString()
    cn.Close
    ComprobarConexion = True
    Exit Function
ErrorHandler:
    ComprobarConexion = False
    If Not cn Is Nothing Then Set cn = Nothing
End Function

Public Function ExecuteNonQuery(ByVal sql As String) As Boolean
    Dim cn As Object
    On Error GoTo ErrorHandler
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()
    cn.Execute sql
    cn.Close
    Set cn = Nothing
    ExecuteNonQuery = True
    Exit Function
ErrorHandler:
    ExecuteNonQuery = False
    If Not cn Is Nothing Then
        If cn.State = 1 Then cn.Close
        Set cn = Nothing
    End If
End Function
```

### 6.2. Persistencia en Caliente de Fila hacia Access (`SincronizarFilaHaciaAccess`)
```vba
Public Function SincronizarFilaHaciaAccess(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    On Error GoTo ErrorHandler
    If ws Is Nothing Or fila < 2 Then Exit Function

    Dim numOrden As Long
    numOrden = CLng(ws.Cells(fila, 1).Value)
    If numOrden <= 0 Then Exit Function

    Dim resTarget As String
    resTarget = ObtenerClaveResidenciaDesdeHoja(ws.Name)

    Dim cn As Object, rs As Object
    Set cn = CreateObject("ADODB.Connection")
    cn.Open GetConnectionString()

    Dim sql As String
    sql = "SELECT * FROM Ordenes WHERE NumOrden = " & numOrden & " AND Residencia = '" & resTarget & "'"
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open sql, cn, 1, 3 ' adOpenKeyset, adLockOptimistic

    If rs.EOF Then
        rs.AddNew
        rs("NumOrden").Value = numOrden
        rs("Residencia").Value = resTarget
        rs("FechaCreacion").Value = Now
        rs("UsuarioCreacion").Value = m_NombreUsuario
    End If

    ' Asignar campos mapeados según residencia
    AsignarCampoFecha rs, "FechaPeticion", ws.Cells(fila, 2).Value
    AsignarCampoNumero rs, "NumFactura", ws.Cells(fila, 3).Value
    AsignarCampoTexto rs, "Finalidad", ws.Cells(fila, 4).Value
    AsignarCampoTexto rs, "Empleo", ws.Cells(fila, 5).Value
    AsignarCampoTexto rs, "DNI", ws.Cells(fila, 9).Value
    AsignarCampoTexto rs, "Nombre", ws.Cells(fila, 11).Value
    AsignarCampoFecha rs, "FechaEntrada", ws.Cells(fila, 12).Value
    AsignarCampoFecha rs, "FechaSalida", ws.Cells(fila, 13).Value
    AsignarCampoNumero rs, "DiasUso", ws.Cells(fila, 14).Value
    AsignarCampoNumero rs, "PAX", ws.Cells(fila, 15).Value
    AsignarCampoDouble rs, "Importe", ws.Cells(fila, 21).Value
    AsignarCampoTexto rs, "EstadoPago", ws.Cells(fila, 27).Value

    rs.Update
    rs.Close
    cn.Close
    Set rs = Nothing
    Set cn = Nothing
    SincronizarFilaHaciaAccess = True
    Exit Function

ErrorHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not cn Is Nothing Then Set cn = Nothing
    SincronizarFilaHaciaAccess = False
End Function
```

### 6.3. Auto-Fit Dinámico de Carátula en Pantalla (`ThisWorkbook.cls`)
```vba
Private Sub AjustarCaratulaInicio(wsInicio As Worksheet)
    On Error Resume Next
    Dim shp As Shape
    Set shp = wsInicio.Shapes("ImgCaratulaCasaCuartel")
    If shp Is Nothing Then Exit Sub

    Dim screenW As Double, screenH As Double
    screenW = Application.UsableWidth
    screenH = Application.UsableHeight
    If screenW <= 0 Or screenH <= 0 Then Exit Sub

    ' Proporción nativa de la imagen (1570 x 845 px = 1.858)
    Const RATIO_FOTO As Double = 1.857988
    Dim maxW As Double, maxH As Double
    maxW = screenW * 0.96
    maxH = screenH * 0.92

    Dim factorW As Double, factorH As Double
    factorW = maxW / RATIO_FOTO
    factorH = maxH

    Dim finalW As Double, finalH As Double
    If factorW <= factorH Then
        finalW = maxW
        finalH = maxW / RATIO_FOTO
    Else
        finalH = maxH
        finalW = maxH * RATIO_FOTO
    End If

    ' Desproteger, redimensionar, centrar y reproteger
    wsInicio.Unprotect Password:=ModuloConfigSegura.ObtenerPasswordHojas()
    shp.LockAspectRatio = msoTrue
    shp.Width = finalW
    shp.Height = finalH
    shp.Left = (screenW - finalW) / 2
    shp.Top = (screenH - finalH) / 2
    wsInicio.Protect Password:=ModuloConfigSegura.ObtenerPasswordHojas(), UserInterfaceOnly:=True
End Sub
```

---

## 7. CATÁLOGO DE RETOS Y PROPUESTAS DE MEJORA ACTIVAS

El repositorio mantiene un catálogo formal de propuestas técnicas para evolución del sistema:

### [PROP-01] Replicación de Altas y Bajas de Lista Negra en Hojas LOG
- **Estado:** 📋 Resuelta / Integrada en sincronización central.
- **Alcance:** Asegurar que cualquier veto introducido en Access se refleje en los registros de auditoría de las tres sedes.

### [PROP-02] Mitigación de Parpadeos Visuales (Flicker) y Alternativa de Splash Screen Nativo
- **Estado:** ⏳ Pendiente de Revisión.
- **Problema:** Al abrir Excel, las imágenes insertadas como `Shape` en cuadrícula sufren repintado en bandas por parte del motor gráfico de Office GDI+/DirectX cuando se superpone un formulario modal.
- **Líneas de Solución:** Valorar migrar la carátula a un formulario UserForm a pantalla completa (Splash Screen nativo) que utiliza búfer de ventana dedicada de Windows sin parpadeo.

### [PROP-03] Sistema de Continuidad Operativa Fuera de Línea (Modo Contingencia Nivel 2 — "Buffer & Sync" / SICO-N2)
- **Estado:** 📋 **Planificada y Priorizada para Próxima Fase.**
- **Problema:** Si la red local o la unidad `H:\` se caen mientras el operador trabaja, o si el PC se enciende sin conexión a primera hora:
  * La versión actual bloquea el acceso en el arranque para evitar trabajar a ciegas.
  * Si se corta sobre la marcha, los cambios se ven en la cuadrícula pero no suben a Access.
- **Solución Diseñada (SICO-N2):**
  1. **Caché Local de Credenciales:** Permite identificarse con la clave habitual aunque no haya red.
  2. **Hoja Oculta `SYS_COLA_SYNC`:** Los cobros en recepción (metálico/banco), cambios de teléfono y notas se graban en una cola no volátil en el archivo `.xlsm`.
  3. **Persistencia Multidiaria:** Si la red tarda días o un fin de semana en volver, los datos persisten intactos tras apagar y reiniciar el PC.
  4. **Semáforo de Conectividad:** 🟢 En línea / 🟡 Modo Contingencia / 🔄 Sincronizando.
  5. **Motor de Reconciliación FIFO:** Al volver la línea, vuelca automáticamente todos los cambios a Access preservando la fecha y hora original exacta de cada cobro.
  6. **Blindaje de Líneas Rojas:** Bloqueo preventivo de numeraciones de factura oficiales y nuevas habitaciones simultáneas para impedir solapes entre sedes.

---

## 8. LÍNEAS DE DEBATE Y PREGUNTAS ABIERTAS PARA GEMINI WEB

Cuando abras este dossier en Gemini Web, puedes plantearle las siguientes líneas estratégicas de consulta:

### Pregunta 1: Arquitectura y Resiliencia (Proyecto SICO-N2)
> *"Revisa la especificación del Proyecto SICO-N2 (Sección 7, PROP-03). ¿Qué mecanismos adicionales de control de concurrencia o verificación de integridad recomendarías para la cola `SYS_COLA_SYNC` para evitar cualquier riesgo de corrupción si el ordenador sufre un corte de suministro eléctrico repentino mientras guardaba en la cola local?"*

### Pregunta 2: Modernización Visual y Experiencia de Usuario (UI/UX)
> *"Teniendo en cuenta que el Front-End es Excel 365/2021 con Ribbon personalizado en XML y UserForms VBA, ¿cuáles son las mejores prácticas modernas para dotar a los formularios de login, selectores y avisos de una apariencia contemporánea y fluida (estilo tarjeta Fluent/Material Design), minimizando los parpadeos de repintado de pantalla?"*

### Pregunta 3: Optimización del Rendimiento en Red Local (LAN/Access)
> *"Con 10 a 15 operadores concurrentes atacando la misma base de datos `Residencia_BE.accdb` en un servidor de archivos compartido, ¿qué ajustes en los índices de Access, timeouts de ADO o consultas SQL recomendarías para minimizar bloqueos de lectura/escritura y maximizar la velocidad de respuesta?"*

### Pregunta 4: Transición Hacia la Versión Web
> *"El proyecto tiene en su horizonte una aplicación web moderna (React + .NET 8). ¿Cómo plantearías una estrategia de convivencia o sincronización híbrida para que los operadores de recepción puedan seguir usando Excel en local mientras ciertas consultas de disponibilidad o reservas se gestionan desde la web?"*

---
*(Fin del Dossier Técnico Maestro — ResidenciaApp 2026)*
