using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.IdentityModel.Tokens;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Application.Options;
using ResidenciaApp.Application.Services;
using ResidenciaApp.Domain.Entities;
using System.Text;

namespace ResidenciaApp.Infrastructure;

public static class DependencyInjection
{
    public static IServiceCollection AddInfrastructure(this IServiceCollection services, IConfiguration configuration)
    {
        var dbPath = Path.Combine(Directory.GetCurrentDirectory(), "residencia.db");
        var connStr = configuration.GetConnectionString("DefaultConnection") ?? $"Data Source={dbPath}";
        var provider = configuration["DatabaseProvider"];

        bool useSqlite = string.Equals(provider, "Sqlite", StringComparison.OrdinalIgnoreCase)
            || connStr.Contains(".db")
            || (connStr.Contains("Data Source=") && !connStr.Contains("Server="))
            || File.Exists(dbPath);

        services.AddDbContext<ResidenciaDbContext>(options =>
        {
            if (useSqlite)
            {
                options.UseSqlite($"Data Source={dbPath}");
            }
            else
            {
                options.UseSqlServer(connStr);
            }
        });

        services.AddScoped<IResidenciaDbContext>(provider => provider.GetRequiredService<ResidenciaDbContext>());

        services.AddIdentity<ApplicationUser, ApplicationRole>(options =>
            {
                options.Password.RequireDigit = true;
                options.Password.RequiredLength = 8;
                options.Password.RequireUppercase = true;
                options.Password.RequireLowercase = true;
                options.Password.RequireNonAlphanumeric = true;
                options.User.RequireUniqueEmail = true;
            })
            .AddEntityFrameworkStores<ResidenciaDbContext>()
            .AddDefaultTokenProviders();

        services.Configure<JwtSettings>(configuration.GetSection("JwtSettings"));

        var jwtSection = configuration.GetSection("JwtSettings");
        var jwtSettings = new JwtSettings
        {
            Secret = jwtSection["Secret"] ?? string.Empty,
            Issuer = jwtSection["Issuer"] ?? string.Empty,
            Audience = jwtSection["Audience"] ?? string.Empty,
            ExpirationMinutes = int.TryParse(jwtSection["ExpirationMinutes"], out var minutes) ? minutes : 60
        };

        bool isProduction = string.Equals(Environment.GetEnvironmentVariable("ASPNETCORE_ENVIRONMENT"), "Production", StringComparison.OrdinalIgnoreCase);

        if (isProduction && (string.IsNullOrWhiteSpace(jwtSettings.Secret) || jwtSettings.Secret.Length < 16))
        {
            throw new InvalidOperationException("SEGURIDAD CRÍTICA: En el entorno de producción DEBE configurarse una clave 'JwtSettings:Secret' segura y robusta.");
        }

        var secretStr = string.IsNullOrWhiteSpace(jwtSettings.Secret) || jwtSettings.Secret.Length < 16
            ? "super-secret-key-for-local-development-123456"
            : jwtSettings.Secret;
        var issuerStr = string.IsNullOrWhiteSpace(jwtSettings.Issuer) ? "ResidenciaApp" : jwtSettings.Issuer;
        var audienceStr = string.IsNullOrWhiteSpace(jwtSettings.Audience) ? "ResidenciaAppClient" : jwtSettings.Audience;

        services.AddAuthentication(options =>
        {
            options.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
            options.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
            options.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
        })
            .AddJwtBearer(options =>
            {
                options.TokenValidationParameters = new TokenValidationParameters
                {
                    ValidateIssuer = true,
                    ValidateAudience = true,
                    ValidateLifetime = true,
                    ValidateIssuerSigningKey = true,
                    ValidIssuer = issuerStr,
                    ValidAudience = audienceStr,
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretStr))
                };
            });

        services.AddAuthorization();
        services.AddScoped<IAuthService, AuthService>();
        services.AddScoped<ISettingsService, SettingsService>();

        // Servicios de configuración
        services.AddScoped<IResidenciaService, ResidenciaService>();
        services.AddScoped<ITipoHabitacionService, TipoHabitacionService>();
        services.AddScoped<IHabitacionService, HabitacionService>();
        services.AddScoped<ITarifaService, TarifaService>();
        services.AddScoped<IFestivoService, FestivoService>();

        // Servicios de reservas
        services.AddScoped<IHuespedService, HuespedService>();
        services.AddScoped<IEvaluacionService, EvaluacionService>();
        services.AddScoped<IDisponibilidadService, DisponibilidadService>();
        services.AddScoped<IReservaService, ReservaService>();

        // Servicios de facturación
        services.AddScoped<IFacturaService, FacturaService>();

        return services;
    }
}
