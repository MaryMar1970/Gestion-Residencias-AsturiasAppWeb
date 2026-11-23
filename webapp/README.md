# ResidenciaApp Web — prototipo independiente

Este proyecto es un prototipo web separado de la aplicación operativa Excel/VBA + Access. No constituye una migración aprobada ni sustituye la especificación arquitectónica de la aplicación operativa.

## Tecnología

- API backend: .NET 8 / ASP.NET Core.
- Persistencia: Entity Framework Core, con proveedor SQL Server configurado por defecto y soporte SQLite en ejecución local.
- Frontend: React, TypeScript y Vite.

## Estructura

- `backend/src/ResidenciaApp.Api/`: API HTTP y configuración de arranque.
- `backend/src/ResidenciaApp.Application/`: casos de uso y servicios de aplicación.
- `backend/src/ResidenciaApp.Domain/`: entidades y reglas de dominio.
- `backend/src/ResidenciaApp.Infrastructure/`: persistencia y adaptadores.
- `frontend/`: cliente React/Vite.
- `database/residencia.db`: archivo SQLite local existente; no es el backend Access de producción ni debe considerarse una base de datos oficial.

## Ejecución local

### Backend

```powershell
Set-Location .\backend
dotnet restore
dotnet run --project .\src\ResidenciaApp.Api\ResidenciaApp.Api.csproj
```

### Frontend

En otra terminal:

```powershell
Set-Location .\frontend
npm install
npm run dev
```

La configuración de puertos se obtiene de los perfiles de lanzamiento del backend y de Vite; no se fija aquí un puerto que pueda quedar obsoleto.

## Base de datos

- La cadena `DefaultConnection` de `backend/src/ResidenciaApp.Api/appsettings.json` configura SQL Server local (`localhost\SQLEXPRESS`, base `ResidenciaAppDb`).
- El arranque aplica migraciones EF Core cuando el proveedor es SQL Server; para el proveedor alternativo ejecuta `EnsureCreated`.
- No existe actualmente un archivo `scripts/create-database.sql` en este subproyecto. No seguir instrucciones antiguas que hagan referencia a ese script.
- `database/residencia.db` es un artefacto local; comprobar la cadena de conexión efectiva antes de asumir que la aplicación lo utiliza.

## Seguridad y estado

- El secreto JWT presente en `appsettings.json` es un valor de desarrollo. No utilizarlo en despliegues; configurar secretos externos y una clave segura por entorno.
- Este prototipo no está integrado ni certificado frente al libro Excel o al backend Access de producción.
- No introducir datos reales de huéspedes ni conectar este prototipo a producción sin una autorización y un diseño de integración independientes.