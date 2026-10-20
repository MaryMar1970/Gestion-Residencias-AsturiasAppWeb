# GUÍA DEL PROYECTO RESIDENCIAAPP PARA GEMINI

Este directorio contiene una **copia limpia y completa del código fuente** del proyecto **ResidenciaApp** (Sistema de Gestión de Residencias de Asturias: Gijón, Soto de Luarca y Oviedo), sin archivos binarios de compilación (`bin/`, `obj/`, `dist/`), sin librerías externas (`node_modules/`), sin control de versiones interno (`.git/`) ni copias de respaldo redundantes.

---

## 1. Propósito del Proyecto
ResidenciaApp es un ecosistema para la reserva, control de ocupación, facturación y auditoría de plazas en las residencias militares y casas cuartel de Asturias (sedes de **Gijón**, **Soto de Luarca** y **Oviedo**). Dispone de un cliente de escritorio multiusuario basado en Microsoft Excel con macros VBA y backend relacional Access, y de una aplicación web moderna (React + .NET 8) orientada a la evolución del sistema.

---

## 2. Estructura de Carpetas

```
ResidenciaApp_CodigoFuente/
├── .agents/                        <- Reglas del proyecto, gobernanza técnica y directivas del agente
│   └── AGENTS.md
├── bdas-multiusuario/              <- Aplicación Cliente Excel VBA Multiusuario
│   ├── BDAS_Multiusuario.xlsm      <- Libro de trabajo activo oficial
│   ├── customUI/                   <- Definición XML de la Cinta de Opciones (Ribbon UI)
│   ├── scripts/                    <- Utilidades PowerShell, VBScript y Python para sincronización y compilación
│   ├── vba-modules/                <- CÓDIGO FUENTE VBA EN TEXTO PLANO (144 componentes: .bas, .cls, .frm)
│   └── REGLAS PERMANENTES...txt    <- Normativa técnica y de integridad del aplicativo
├── docs/                           <- Documentación Técnica, Arquitectura y Esquemas
│   ├── access_schema.txt           <- Esquema relacional de tablas y campos de Access (Residencia_BE.accdb)
│   ├── architecture_summary.json   <- Mapa estructurado de módulos, hojas y dependencias
│   ├── sheet_details.json          <- Detalle de las 50 hojas Excel (CodeNames, nombres visibles, propósitos)
│   ├── MAPA_DEPENDENCIAS.md        <- Diagrama de dependencias entre módulos y formularios
│   ├── DICCIONARIO_RANGOS.md       <- Nombres definidos y rangos clave de cada hoja
│   ├── INDICE_VERSION_REFERENCIA...<- Hitos de versiones de referencia oficiales
│   └── REGISTRO_CAMBIOS...         <- Histórico detallado de intervenciones y soluciones aplicadas
├── webapp/                         <- Aplicación Web Moderna (Full Stack)
│   ├── backend/                    <- API REST en C# .NET 8 (Arquitectura Limpia: Api, Application, Domain, Infrastructure)
│   │   ├── ResidenciaApp.sln
│   │   └── src/
│   └── frontend/                   <- Single Page Application en React + TypeScript + Vite
│       ├── src/
│       ├── index.html
│       ├── package.json
│       ├── tsconfig.json
│       └── vite.config.ts
├── README.md                       <- Información general del repositorio
├── .editorconfig
├── .gitattributes
└── .gitignore
```

---

## 3. Componentes Clave por Tecnología

### A. Cliente Excel / VBA (`bdas-multiusuario/vba-modules/`)
Todo el código VBA ha sido exportado en texto plano para facilitar su análisis, búsqueda e interpretación por LLMs:
- **`modDatabase.bas`**: Capa de persistencia relacional con Access vía ADODB (UPSERT, consultas seguras, reconexión automática).
- **`modSeguridad.bas`**: Control de acceso, sesiones persistentes en nombres definidos ocultos y permisos por sede/rol.
- **`modAuditoria.bas`**: Registro transaccional estricto en hojas LOG y tabla remota `Auditoria`.
- **`ModuloCalendarioGijon.bas`, `ModuloCalendarioSoto.bas`, `ModuloCalendarioOviedo.bas`**: Motores de calendario, turnos y validación de ocupación mensual.
- **Hojas de Residencia (`Hoja1.cls` a `Hoja50.cls`)**: Eventos de hoja (`Worksheet_Change`, `Worksheet_Calculate`) y validación de celdas.
- **Formularios (`frmLogin.frm`, `frmSelectorResidencia.frm`, `frmGestionListaNegra.frm`)**: Interfaces visuales de usuario.

### B. Base de Datos Relacional (`docs/access_schema.txt`)
La base de datos Access backend (`Residencia_BE.accdb`) contiene las tablas maestras:
- `Ordenes` y `Ordenes_Historico`: Registro de adjudicaciones, ocupantes, habitaciones y períodos.
- `Usuarios` y `Usuarios_Sedes`: Control de credenciales, roles (ADMIN, OPERADOR, LECTURA) y residencias asignadas.
- `Auditoria`: Trazabilidad completa de operaciones.
- `Precios` / `Tarifas`: Desglose económico de alojamientos y suplementos.

### C. Aplicación Web (`webapp/`)
- **Backend (.NET 8)**: Estructurado según *Clean Architecture* / *DDD*, con Entity Framework Core, repositorios y controladores REST.
- **Frontend (React + Vite)**: Componentes modulares tipados en TypeScript con gestión de estado y diseño responsivo para consulta y reserva de plazas.

---

## 4. Reglas Críticas del Proyecto (`.agents/AGENTS.md`)
1. **Regla de Nombres**: Se utiliza estrictamente `GIJÓN` (o `GIJON`). Prohibido el uso de concatenaciones de caracteres tipo `Chr(211)`.
2. **Protección contra Inyección de Fórmulas**: Los datos de auditoría que inician con signo igual (`=`) se prefijan con apóstrofo para neutralizar el Error 1004.
3. **Resistencia de Sesión**: La sesión de usuario se mantiene persistente en Nombres Definidos Ocultos de Excel, evitando pérdidas por reseteo de variables globales.
4. **Borrado Físico RGPD**: El borrado definitivo de datos personales de la base de datos está restringido de forma exclusiva a la interfaz designada para RGPD.
