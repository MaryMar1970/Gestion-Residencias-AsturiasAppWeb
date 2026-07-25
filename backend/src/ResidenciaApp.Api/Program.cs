using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Domain.Entities;
using ResidenciaApp.Infrastructure;

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
    dbContext.Database.Migrate();

    // Healing routine for reservation order numbers (assign sequential numbers to existing guest reservations where it is 0)
    var reservationsToHeal = dbContext.Reservas
        .Where(r => r.NumeroOrden == 0 && r.HuespedId != null && !r.EsBloqueo)
        .OrderBy(r => r.FechaEntrada)
        .ToList();

    if (reservationsToHeal.Any())
    {
        int currentMax = dbContext.Reservas
            .Where(r => r.NumeroOrden > 0 && r.HuespedId != null && !r.EsBloqueo)
            .Select(r => (int?)r.NumeroOrden)
            .Max() ?? 0;

        foreach (var r in reservationsToHeal)
        {
            currentMax++;
            r.NumeroOrden = currentMax;
        }
        dbContext.SaveChanges();
    }

    // Healing routine for empty or null Resolucion values (set to "SI" by default)
    var reservationsWithoutResolution = dbContext.Reservas
        .Where(r => r.Resolucion == null || r.Resolucion == "")
        .ToList();

    if (reservationsWithoutResolution.Any())
    {
        foreach (var r in reservationsWithoutResolution)
        {
            r.Resolucion = "SI";
        }
        dbContext.SaveChanges();
    }

    await SeedAsync(services);
}

app.Run();

static async Task SeedAsync(IServiceProvider services)
{
    var roleManager = services.GetRequiredService<RoleManager<ApplicationRole>>();
    var userManager = services.GetRequiredService<UserManager<ApplicationUser>>();
    var context = services.GetRequiredService<ResidenciaDbContext>();

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

    var requiredSettings = new[]
    {
        new SystemSetting { Key = "ResidenceName", Value = "Residencia Demo", Description = "Nombre de la residencia", Category = "General" },
        new SystemSetting { Key = "RoomCount", Value = "12", Description = "Número total de habitaciones", Category = "Habitaciones" },
        new SystemSetting { Key = "RoomTypes", Value = "Simple,Doble,Suite", Description = "Tipos de habitaciones disponibles", Category = "Habitaciones" },
        new SystemSetting { Key = "DailyRate", Value = "95", Description = "Precio diario base", Category = "Precios" },
        new SystemSetting { Key = "TaxRate", Value = "0.21", Description = "IVA aplicado", Category = "Precios" },
        new SystemSetting { Key = "BlockedDates", Value = "", Description = "Bloqueos de fechas", Category = "Reserva" },
        new SystemSetting { Key = "AvoidOverlaps", Value = "true", Description = "Evitar solapes en reservas", Category = "Reserva" }
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
        context.Residencias.Add(new Residencia
        {
            Nombre = "Residencia Test Python",
            Activa = true,
            Orden = 1
        });
        context.Residencias.Add(new Residencia
        {
            Nombre = "Segunda Residencia",
            Activa = true,
            Orden = 2
        });
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
