using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Domain.Entities;
using ResidenciaApp.Infrastructure;

// Garantizar que la DLL nativa e_sqlite3.dll esté en la raíz de ejecución para evitar el límite MAX_PATH de Windows (Error 0x800700CE / bad_module_info)
try
{
    var baseDir = AppDomain.CurrentDomain.BaseDirectory;
    var targetFile = Path.Combine(baseDir, "e_sqlite3.dll");
    if (!File.Exists(targetFile))
    {
        var candidate = Path.Combine(baseDir, "runtimes", "win-x64", "native", "e_sqlite3.dll");
        if (File.Exists(candidate))
        {
            File.Copy(candidate, targetFile, true);
        }
    }
}
catch { /* ignore */ }

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowFrontend", policy =>
    {
        policy.SetIsOriginAllowed(origin =>
                new Uri(origin).Host == "localhost")
            .AllowAnyHeader()
            .AllowAnyMethod();
    });
});

builder.Services.AddControllers();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

builder.Services.AddInfrastructure(builder.Configuration);

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseCors("AllowFrontend");
app.UseAuthentication();
app.UseAuthorization();
app.MapControllers();

using (var scope = app.Services.CreateScope())
{
    var services = scope.ServiceProvider;
    var dbContext = services.GetRequiredService<ResidenciaDbContext>();
    var logger = services.GetRequiredService<ILogger<Program>>();

    try
    {
        if (dbContext.Database.IsSqlServer())
        {
            dbContext.Database.Migrate();
        }
        else
        {
            dbContext.Database.EnsureCreated();
        }
    }
    catch (Exception ex)
    {
        logger.LogWarning(ex, "Advertencia durante la preparación o migración de la base de datos.");
    }



    try
    {
        await SeedAsync(services);
    }
    catch (Exception ex)
    {
        logger.LogWarning(ex, "Advertencia durante el sembrado de datos de prueba.");
    }
}

app.Run();

static async Task SeedAsync(IServiceProvider services)
{
    var context = services.GetRequiredService<ResidenciaDbContext>();
    try
    {
        await context.Database.MigrateAsync();
    }
    catch
    {
        await context.Database.EnsureCreatedAsync();
    }

    try
    {
        await context.Database.ExecuteSqlRawAsync(@"
            IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Residencias')
            BEGIN
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Residencias' AND COLUMN_NAME = 'SerieFactura')
                    ALTER TABLE [Residencias] ADD [SerieFactura] nvarchar(20) NULL;
            END
            IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Habitaciones')
            BEGIN
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Habitaciones' AND COLUMN_NAME = 'TipoCamaPrincipal')
                    ALTER TABLE [Habitaciones] ADD [TipoCamaPrincipal] nvarchar(50) NULL;
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Habitaciones' AND COLUMN_NAME = 'CapacidadPersonas')
                    ALTER TABLE [Habitaciones] ADD [CapacidadPersonas] int NULL;
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Habitaciones' AND COLUMN_NAME = 'AdmiteSupletorias')
                    ALTER TABLE [Habitaciones] ADD [AdmiteSupletorias] bit NULL;
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Habitaciones' AND COLUMN_NAME = 'PlazasSupletorias')
                    ALTER TABLE [Habitaciones] ADD [PlazasSupletorias] int NULL;
            END
            IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'Reservas')
            BEGIN
                IF EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Reservas' AND COLUMN_NAME = 'HabitacionId' AND IS_NULLABLE = 'NO')
                    ALTER TABLE [Reservas] ALTER COLUMN [HabitacionId] uniqueidentifier NULL;
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Reservas' AND COLUMN_NAME = 'FechaSolicitud')
                    ALTER TABLE [Reservas] ADD [FechaSolicitud] datetime2 NULL;
                IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Reservas' AND COLUMN_NAME = 'ResidenciaId')
                BEGIN
                    ALTER TABLE [Reservas] ADD [ResidenciaId] uniqueidentifier NULL;
                    EXEC('UPDATE [Reservas] SET [ResidenciaId] = (SELECT [ResidenciaId] FROM [Habitaciones] WHERE [Habitaciones].[Id] = [Reservas].[HabitacionId]) WHERE [ResidenciaId] IS NULL AND [HabitacionId] IS NOT NULL;');
                END
            END
        ");
    }
    catch { /* Ignore if unsupported DB engine */ }

    var roleManager = services.GetRequiredService<RoleManager<ApplicationRole>>();
    var userManager = services.GetRequiredService<UserManager<ApplicationUser>>();

    if (!await roleManager.RoleExistsAsync("Admin"))
    {
        await roleManager.CreateAsync(new ApplicationRole { Name = "Admin", NormalizedName = "ADMIN" });
    }

    var admin = await userManager.FindByEmailAsync("admin@residencia.local");
    if (admin is null)
    {
        admin = new ApplicationUser
        {
            UserName = "admin@residencia.local",
            Email = "admin@residencia.local",
            FullName = "Administrador",
            EmailConfirmed = true
        };

        var result = await userManager.CreateAsync(admin, "Admin123!");
        if (result.Succeeded)
        {
            await userManager.AddToRoleAsync(admin, "Admin");
        }
    }

    var defaultMatrizJson = System.Text.Json.JsonSerializer.Serialize(ResidenciaApp.Application.Services.EvaluacionService.MatrizPorDefecto);

    var requiredSettings = new[]
    {
        new SystemSetting { Key = "ResidenceName", Value = "Residencia Demo", Description = "Nombre de la residencia", Category = "General" },
        new SystemSetting { Key = "RoomCount", Value = "12", Description = "Número total de habitaciones", Category = "Habitaciones" },
        new SystemSetting { Key = "RoomTypes", Value = "Simple,Doble,Suite", Description = "Tipos de habitaciones disponibles", Category = "Habitaciones" },
        new SystemSetting { Key = "DailyRate", Value = "95", Description = "Precio diario base", Category = "Precios" },
        new SystemSetting { Key = "TaxRate", Value = "0.21", Description = "IVA aplicado", Category = "Precios" },
        new SystemSetting { Key = "BlockedDates", Value = "", Description = "Bloqueos de fechas", Category = "Reserva" },
        new SystemSetting { Key = "AvoidOverlaps", Value = "true", Description = "Evitar solapes en reservas", Category = "Reserva" },
        new SystemSetting { Key = "MatrizEvaluacion", Value = defaultMatrizJson, Description = "Matriz de prioridades y evaluación de solicitudes", Category = "Evaluacion" }
    };

    foreach (var setting in requiredSettings)
    {
        if (!await context.SystemSettings.AnyAsync(x => x.Key == setting.Key))
        {
            context.SystemSettings.Add(setting);
        }
    }

    if (!await context.Residencias.AnyAsync())
    {
        var resGijon = new Residencia { Nombre = "Residencia Gijón", SerieFactura = "GIJ", Cif = "Q3300001A", Direccion = "Av. del Portugete 12", Municipio = "Gijón", Provincia = "Asturias", Activa = true, Orden = 1 };
        var resSoto  = new Residencia { Nombre = "Residencia Soto del Real", SerieFactura = "SOT", Cif = "Q2800002B", Direccion = "Calle Real 45", Municipio = "Soto del Real", Provincia = "Madrid", Activa = true, Orden = 2 };
        var resOviedo = new Residencia { Nombre = "Residencia Oviedo", SerieFactura = "OVI", Cif = "Q3300003C", Direccion = "Calle Uría 8", Municipio = "Oviedo", Provincia = "Asturias", Activa = true, Orden = 3 };

        context.Residencias.AddRange(resGijon, resSoto, resOviedo);
        await context.SaveChangesAsync();

        var tInd = new TipoHabitacion { Nombre = "Individual", Codigo = "IND", CapacidadMaxima = 1, AdmiteSupletorias = false, Descripcion = "Habitación individual estándar", Activo = true, Orden = 1 };
        var tDob = new TipoHabitacion { Nombre = "Doble", Codigo = "DOB", CapacidadMaxima = 2, AdmiteSupletorias = true, Descripcion = "Habitación doble matrimonial o twin", Activo = true, Orden = 2 };
        var tSui = new TipoHabitacion { Nombre = "Suite", Codigo = "SUI", CapacidadMaxima = 4, AdmiteSupletorias = true, Descripcion = "Suite familiar con salón independiente", Activo = true, Orden = 3 };
        var tApt = new TipoHabitacion { Nombre = "Apartamento", Codigo = "APT", CapacidadMaxima = 4, AdmiteSupletorias = true, Descripcion = "Apartamento completo", Activo = true, Orden = 4 };

        context.TiposHabitacion.AddRange(tInd, tDob, tSui, tApt);
        await context.SaveChangesAsync();

        var hab101 = new Habitacion { ResidenciaId = resGijon.Id, TipoHabitacionId = tInd.Id, Numero = "101", Nombre = "Habitación 101", TipoCamaPrincipal = "IND", CapacidadPersonas = 1, AdmiteSupletorias = false, PlazasSupletorias = 0, Activa = true, Orden = 1 };
        var hab102 = new Habitacion { ResidenciaId = resGijon.Id, TipoHabitacionId = tDob.Id, Numero = "102", Nombre = "Habitación 102", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 2 };
        var hab103 = new Habitacion { ResidenciaId = resGijon.Id, TipoHabitacionId = tDob.Id, Numero = "103", Nombre = "Habitación 103", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 3 };
        var hab104 = new Habitacion { ResidenciaId = resGijon.Id, TipoHabitacionId = tSui.Id, Numero = "104", Nombre = "Suite 104", TipoCamaPrincipal = "DOB", CapacidadPersonas = 4, AdmiteSupletorias = true, PlazasSupletorias = 2, Activa = true, Orden = 4 };

        var hab201 = new Habitacion { ResidenciaId = resSoto.Id, TipoHabitacionId = tInd.Id, Numero = "201", Nombre = "Habitación 201", TipoCamaPrincipal = "IND", CapacidadPersonas = 1, AdmiteSupletorias = false, PlazasSupletorias = 0, Activa = true, Orden = 1 };
        var hab202 = new Habitacion { ResidenciaId = resSoto.Id, TipoHabitacionId = tDob.Id, Numero = "202", Nombre = "Habitación 202", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 2 };

        var hab301 = new Habitacion { ResidenciaId = resOviedo.Id, TipoHabitacionId = tDob.Id, Numero = "301", Nombre = "Habitación 301", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 1 };

        context.Habitaciones.AddRange(hab101, hab102, hab103, hab104, hab201, hab202, hab301);
        await context.SaveChangesAsync();

        var h1 = new Huesped { Dni = "12345678A", Nombre = "Carlos", Apellidos = "García Rodríguez", Telefono = "600111222", Email = "carlos.garcia@example.com", TipoHuesped = "Socio", Empleo = "Guardia", Situacion = "Activo", Finalidad = "Comisión", EmpleoCategoria = "GC", FamiliaNumerosa = "NO" };
        var h2 = new Huesped { Dni = "87654321B", Nombre = "María", Apellidos = "López Fernández", Telefono = "611222333", Email = "maria.lopez@example.com", TipoHuesped = "Socio", Empleo = "Cabo", Situacion = "Activo", Finalidad = "Destino", EmpleoCategoria = "GC", FamiliaNumerosa = "GENERAL", PorcentajeDescuento = 20 };
        var h3 = new Huesped { Dni = "45678912C", Nombre = "Javier", Apellidos = "Martínez Sánchez", Telefono = "622333444", Email = "javier.martinez@example.com", TipoHuesped = "Externo", Empleo = "Sargento", Situacion = "Reserva", Finalidad = "Otros", EmpleoCategoria = "GC", FamiliaNumerosa = "NO" };
        var h4 = new Huesped { Dni = "98765432D", Nombre = "Ana", Apellidos = "Gómez Pérez", Telefono = "633444555", Email = "ana.gomez@example.com", TipoHuesped = "Socio", Empleo = "Guardia", Situacion = "Viogen", Finalidad = "Otros", EmpleoCategoria = "GC", FamiliaNumerosa = "ESPECIAL", PorcentajeDescuento = 50 };

        context.Huespedes.AddRange(h1, h2, h3, h4);
        await context.SaveChangesAsync();

        var tarifa1 = new Tarifa { ResidenciaId = resGijon.Id, TipoHabitacionId = tInd.Id, NombreTarifa = "Estándar Individual", PrecioNoche = 45m, PorcentajeIva = 10m, Activa = true };
        var tarifa2 = new Tarifa { ResidenciaId = resGijon.Id, TipoHabitacionId = tDob.Id, NombreTarifa = "Estándar Doble", PrecioNoche = 65m, PorcentajeIva = 10m, Activa = true };
        var tarifa3 = new Tarifa { ResidenciaId = resGijon.Id, TipoHabitacionId = tSui.Id, NombreTarifa = "Estándar Suite", PrecioNoche = 100m, PorcentajeIva = 10m, Activa = true };

        context.Tarifas.AddRange(tarifa1, tarifa2, tarifa3);
        await context.SaveChangesAsync();

        var r1 = new Reserva
        {
            HabitacionId = hab101.Id,
            HuespedId = h1.Id,
            NumeroOrden = 1,
            FechaEntrada = new DateOnly(2026, 8, 1),
            FechaSalida = new DateOnly(2026, 8, 10),
            NumPersonas = 1, NumNinos = 0, CamasSupletorias = 0,
            FamiliaNumerosa = "NO", PorcentajeDescuento = 0,
            TarifaId = tarifa1.Id, TarifaNombreSnapshot = tarifa1.NombreTarifa,
            PrecioNocheAplicado = 45m, PorcentajeIvaAplicado = 10m, TotalNoches = 9,
            ImporteBase = 405m, ImporteIva = 40.5m, ImporteTotal = 445.5m,
            Estado = EstadoReserva.Confirmada, Pagado = false,
            Finalidad = "Comisión", Empleo = "GC", Evaluacion = "1, 1, 4",
            Resolucion = Resoluciones.Si, FechaSolicitud = DateTime.UtcNow.AddDays(-15)
        };

        var r2 = new Reserva
        {
            HabitacionId = hab102.Id,
            HuespedId = h2.Id,
            NumeroOrden = 2,
            FechaEntrada = new DateOnly(2026, 8, 5),
            FechaSalida = new DateOnly(2026, 8, 12),
            NumPersonas = 2, NumNinos = 0, CamasSupletorias = 0,
            FamiliaNumerosa = "GENERAL", PorcentajeDescuento = 20,
            TarifaId = tarifa2.Id, TarifaNombreSnapshot = tarifa2.NombreTarifa,
            PrecioNocheAplicado = 65m, PorcentajeIvaAplicado = 10m, TotalNoches = 7,
            ImporteBase = 364m, ImporteIva = 36.4m, ImporteTotal = 400.4m,
            Estado = EstadoReserva.Confirmada, Pagado = true, FormaPago = "Tarjeta", FechaPago = DateTime.UtcNow.AddDays(-2),
            Finalidad = "Destino", Empleo = "GC", Evaluacion = "1, 1, 3",
            Resolucion = Resoluciones.Concedida, FechaSolicitud = DateTime.UtcNow.AddDays(-10)
        };

        var r3 = new Reserva
        {
            HabitacionId = hab104.Id,
            HuespedId = h4.Id,
            NumeroOrden = 3,
            FechaEntrada = new DateOnly(2026, 8, 2),
            FechaSalida = new DateOnly(2026, 8, 15),
            NumPersonas = 2, NumNinos = 1, CamasSupletorias = 1,
            FamiliaNumerosa = "ESPECIAL", PorcentajeDescuento = 50,
            TarifaId = tarifa3.Id, TarifaNombreSnapshot = tarifa3.NombreTarifa,
            PrecioNocheAplicado = 100m, PorcentajeIvaAplicado = 10m, TotalNoches = 13,
            ImporteBase = 650m, ImporteIva = 65m, ImporteTotal = 715m,
            Estado = EstadoReserva.Confirmada, Pagado = false,
            Finalidad = "Otros", Empleo = "GC", Evaluacion = "2, 1, 2",
            Resolucion = Resoluciones.Concedida, FechaSolicitud = DateTime.UtcNow.AddDays(-12)
        };

        var r4 = new Reserva
        {
            HabitacionId = hab202.Id,
            HuespedId = h3.Id,
            NumeroOrden = 4,
            FechaEntrada = new DateOnly(2026, 8, 1),
            FechaSalida = new DateOnly(2026, 8, 8),
            NumPersonas = 2, NumNinos = 0, CamasSupletorias = 0,
            FamiliaNumerosa = "NO", PorcentajeDescuento = 0,
            PrecioNocheAplicado = 60m, PorcentajeIvaAplicado = 10m, TotalNoches = 7,
            ImporteBase = 420m, ImporteIva = 42m, ImporteTotal = 462m,
            Estado = EstadoReserva.Confirmada, Pagado = false,
            Finalidad = "Otros", Empleo = "GC", Evaluacion = "2, 3, 7",
            Resolucion = Resoluciones.Si, FechaSolicitud = DateTime.UtcNow.AddDays(-8)
        };

        context.Reservas.AddRange(r1, r2, r3, r4);
        await context.SaveChangesAsync();
    }

    if (!await context.Habitaciones.AnyAsync())
    {
        var residenciasList = await context.Residencias.ToListAsync();
        var tInd = await context.TiposHabitacion.FirstOrDefaultAsync(t => t.Codigo == "IND") ?? await context.TiposHabitacion.FirstAsync();
        var tDob = await context.TiposHabitacion.FirstOrDefaultAsync(t => t.Codigo == "DOB") ?? await context.TiposHabitacion.FirstAsync();
        var tSui = await context.TiposHabitacion.FirstOrDefaultAsync(t => t.Codigo == "SUI") ?? await context.TiposHabitacion.FirstAsync();

        foreach (var res in residenciasList)
        {
            var h101 = new Habitacion { ResidenciaId = res.Id, TipoHabitacionId = tInd.Id, Numero = "101", Nombre = "Habitación 101", TipoCamaPrincipal = "IND", CapacidadPersonas = 1, AdmiteSupletorias = false, PlazasSupletorias = 0, Activa = true, Orden = 1 };
            var h102 = new Habitacion { ResidenciaId = res.Id, TipoHabitacionId = tDob.Id, Numero = "102", Nombre = "Habitación 102", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 2 };
            var h103 = new Habitacion { ResidenciaId = res.Id, TipoHabitacionId = tDob.Id, Numero = "103", Nombre = "Habitación 103", TipoCamaPrincipal = "DOB", CapacidadPersonas = 2, AdmiteSupletorias = true, PlazasSupletorias = 1, Activa = true, Orden = 3 };
            var h104 = new Habitacion { ResidenciaId = res.Id, TipoHabitacionId = tSui.Id, Numero = "104", Nombre = "Suite 104", TipoCamaPrincipal = "DOB", CapacidadPersonas = 4, AdmiteSupletorias = true, PlazasSupletorias = 2, Activa = true, Orden = 4 };
            context.Habitaciones.AddRange(h101, h102, h103, h104);
        }
        await context.SaveChangesAsync();
    }

    if (!await context.CodigosPostales.AnyAsync())
    {
        var csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "codigos_postales.csv");
        if (!File.Exists(csvPath))
        {
            csvPath = Path.Combine(Directory.GetCurrentDirectory(), "..", "ResidenciaApp.Infrastructure", "codigos_postales.csv");
            if (!File.Exists(csvPath))
            {
                csvPath = Path.Combine(Directory.GetCurrentDirectory(), "codigos_postales.csv");
            }
        }

        if (File.Exists(csvPath))
        {
            var lines = await File.ReadAllLinesAsync(csvPath);
            var list = new List<CodigoPostalInfo>();
            foreach (var line in lines)
            {
                var parts = line.Split(';');
                if (parts.Length >= 3)
                {
                    list.Add(new CodigoPostalInfo
                    {
                        CodigoPostal = parts[0].Trim(),
                        Municipio = parts[1].Trim(),
                        Provincia = parts[2].Trim()
                    });
                }
            }

            if (list.Any())
            {
                const int chunkSize = 2000;
                for (int i = 0; i < list.Count; i += chunkSize)
                {
                    var chunk = list.Skip(i).Take(chunkSize);
                    context.CodigosPostales.AddRange(chunk);
                    await context.SaveChangesAsync();
                }
            }
        }
    }

    await context.SaveChangesAsync();
}
