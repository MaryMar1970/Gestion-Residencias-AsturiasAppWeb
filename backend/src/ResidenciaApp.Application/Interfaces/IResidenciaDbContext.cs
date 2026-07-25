using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Interfaces;

public interface IResidenciaDbContext
{
    DbSet<SystemSetting> SystemSettings { get; }
    DbSet<UserSetting> UserSettings { get; }
    DbSet<Residencia> Residencias { get; }
    DbSet<TipoHabitacion> TiposHabitacion { get; }
    DbSet<Habitacion> Habitaciones { get; }
    DbSet<Tarifa> Tarifas { get; }
    DbSet<Festivo> Festivos { get; }
    DbSet<Huesped> Huespedes { get; }
    DbSet<Reserva> Reservas { get; }
    DbSet<Factura> Facturas { get; }
    DbSet<LineaFactura> LineasFactura { get; }
    DbSet<CodigoPostalInfo> CodigosPostales { get; }

    Task<int> SaveChangesAsync(CancellationToken cancellationToken = default);
}
