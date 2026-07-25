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
        services.AddDbContext<ResidenciaDbContext>(options =>
            options.UseSqlServer(configuration.GetConnectionString("DefaultConnection")));

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
                    ValidIssuer = jwtSettings.Issuer,
                    ValidAudience = jwtSettings.Audience,
                    IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtSettings.Secret))
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
        services.AddScoped<IReservaService, ReservaService>();

        // Servicios de facturación
        services.AddScoped<IFacturaService, FacturaService>();

        return services;
    }
}
