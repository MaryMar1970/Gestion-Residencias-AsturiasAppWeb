# Residencia App

Base de referencia para una aplicación web multiusuario orientada a la gestión de una residencia, con:

- Arquitectura en capas: API, Application, Infrastructure y Domain
- Autenticación JWT + ASP.NET Identity
- Entity Framework Core + SQL Server
- Frontend React + TypeScript con Vite

## Estructura

- backend/src/ResidenciaApp.Api: API web ASP.NET Core
- backend/src/ResidenciaApp.Application: lógica de aplicación y DTOs
- backend/src/ResidenciaApp.Infrastructure: EF Core, Identity y servicios de infraestructura
- backend/src/ResidenciaApp.Domain: entidades del dominio
- frontend: aplicación React con pantalla de login y gestión de parámetros
- scripts/create-database.sql: script SQL Server para crear la base de datos

## Requisitos locales

- .NET 8 SDK
- Node.js 20+
- SQL Server (LocalDB, SQL Express o SQL Server Developer)

## 1) Base de datos

Ejecuta el script SQL Server desde [scripts/create-database.sql](scripts/create-database.sql) o ajusta la cadena de conexión en [backend/src/ResidenciaApp.Api/appsettings.json](backend/src/ResidenciaApp.Api/appsettings.json).

Ejemplo de cadena de conexión:

```json
"ConnectionStrings": {
  "DefaultConnection": "Server=(localdb)\\MSSQLLocalDB;Database=ResidenciaAppDb;Trusted_Connection=True;TrustServerCertificate=True;"
}
```

Si utilizas SQL Server Express o una instancia con nombre, ajusta la cadena de conexión a algo como:

```json
"ConnectionStrings": {
  "DefaultConnection": "Server=.\\SQLEXPRESS;Database=ResidenciaAppDb;Trusted_Connection=True;TrustServerCertificate=True;"
}
```

## 2) Backend

```bash
cd backend
dotnet restore
dotnet run --project src/ResidenciaApp.Api/ResidenciaApp.Api.csproj
```

La API quedará disponible en http://localhost:5260.

Endpoints incluidos:

- POST /auth/login
- GET /users/me
- GET /settings
- POST /settings

## 3) Frontend

```bash
cd frontend
npm install
npm run dev
```

La app quedará disponible en http://localhost:5173.

## Credenciales iniciales

- Email: admin@residencia.local
- Password: Admin123!

## Próximos pasos recomendados

- Añadir más entidades de negocio (habitaciones, reservas, contratos, clientes)
- Extender los parámetros a un modelo de configuración más rico
- Añadir permisos por rol para operaciones de administración y reservas
- Sustituir la gestión de parámetros por un modelo más escalable de “configuraciones por residencia”
