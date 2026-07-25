using Microsoft.AspNetCore.Identity.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Infrastructure;

public class ResidenciaDbContext : IdentityDbContext<ApplicationUser, ApplicationRole, Guid>, IResidenciaDbContext
{
    public ResidenciaDbContext(DbContextOptions<ResidenciaDbContext> options)
        : base(options)
    {
    }

    public DbSet<SystemSetting> SystemSettings => Set<SystemSetting>();
    public DbSet<UserSetting> UserSettings => Set<UserSetting>();
    public DbSet<Residencia> Residencias => Set<Residencia>();
    public DbSet<TipoHabitacion> TiposHabitacion => Set<TipoHabitacion>();
    public DbSet<Habitacion> Habitaciones => Set<Habitacion>();
    public DbSet<Tarifa> Tarifas => Set<Tarifa>();
    public DbSet<Festivo> Festivos => Set<Festivo>();
    public DbSet<Huesped> Huespedes => Set<Huesped>();
    public DbSet<Reserva> Reservas => Set<Reserva>();
    public DbSet<Factura> Facturas => Set<Factura>();
    public DbSet<LineaFactura> LineasFactura => Set<LineaFactura>();
    public DbSet<CodigoPostalInfo> CodigosPostales => Set<CodigoPostalInfo>();

    protected override void OnModelCreating(ModelBuilder builder)
    {
        base.OnModelCreating(builder);

        builder.Entity<SystemSetting>(entity =>
        {
            entity.HasIndex(x => x.Key).IsUnique();
            entity.Property(x => x.Key).IsRequired().HasMaxLength(120);
            entity.Property(x => x.Value).IsRequired().HasMaxLength(4000);
            entity.Property(x => x.Category).IsRequired().HasMaxLength(80);
        });

        builder.Entity<UserSetting>(entity =>
        {
            entity.HasIndex(x => new { x.UserId, x.Key }).IsUnique();
            entity.Property(x => x.Key).IsRequired().HasMaxLength(120);
            entity.Property(x => x.Value).IsRequired().HasMaxLength(4000);
            entity.HasOne(x => x.User)
                .WithMany(x => x.Settings)
                .HasForeignKey(x => x.UserId)
                .OnDelete(DeleteBehavior.Cascade);
        });

        builder.Entity<Residencia>(entity =>
        {
            entity.HasIndex(x => x.Nombre).IsUnique();
            entity.Property(x => x.Nombre).IsRequired().HasMaxLength(120);
            entity.Property(x => x.Cif).HasMaxLength(20);
            entity.Property(x => x.CodigoPostal).HasMaxLength(10);
            entity.Property(x => x.Telefono).HasMaxLength(20);
            entity.Property(x => x.Email).HasMaxLength(200);
        });

        builder.Entity<TipoHabitacion>(entity =>
        {
            entity.HasIndex(x => x.Codigo).IsUnique();
            entity.Property(x => x.Nombre).IsRequired().HasMaxLength(80);
            entity.Property(x => x.Codigo).IsRequired().HasMaxLength(10);
        });

        builder.Entity<Habitacion>(entity =>
        {
            entity.HasIndex(x => new { x.ResidenciaId, x.Numero }).IsUnique();
            entity.Property(x => x.Numero).IsRequired().HasMaxLength(20);
            entity.Property(x => x.Nombre).HasMaxLength(100);
            entity.HasOne(x => x.Residencia)
                .WithMany(x => x.Habitaciones)
                .HasForeignKey(x => x.ResidenciaId)
                .OnDelete(DeleteBehavior.Cascade);
            entity.HasOne(x => x.TipoHabitacion)
                .WithMany(x => x.Habitaciones)
                .HasForeignKey(x => x.TipoHabitacionId)
                .OnDelete(DeleteBehavior.Restrict);
        });

        builder.Entity<Tarifa>(entity =>
        {
            entity.Property(x => x.NombreTarifa).IsRequired().HasMaxLength(80);
            entity.Property(x => x.PrecioNoche).HasColumnType("decimal(10,2)");
            entity.Property(x => x.PrecioMes).HasColumnType("decimal(10,2)");
            entity.Property(x => x.PorcentajeIva).HasColumnType("decimal(5,2)");
            entity.HasOne(x => x.TipoHabitacion)
                .WithMany(x => x.Tarifas)
                .HasForeignKey(x => x.TipoHabitacionId)
                .OnDelete(DeleteBehavior.Cascade);
            entity.HasOne(x => x.Residencia)
                .WithMany(x => x.Tarifas)
                .HasForeignKey(x => x.ResidenciaId)
                .OnDelete(DeleteBehavior.Cascade);
        });

        builder.Entity<Festivo>(entity =>
        {
            entity.HasIndex(x => x.Fecha).IsUnique();
            entity.Property(x => x.Descripcion).IsRequired().HasMaxLength(200);
            entity.Property(x => x.Ambito).HasMaxLength(20);
        });

        builder.Entity<Huesped>(entity =>
        {
            entity.HasIndex(x => x.Dni).IsUnique();
            entity.Property(x => x.Dni).IsRequired().HasMaxLength(20);
            entity.Property(x => x.Nombre).IsRequired().HasMaxLength(100);
            entity.Property(x => x.Apellidos).IsRequired().HasMaxLength(150);
            entity.Property(x => x.Telefono).HasMaxLength(30);
            entity.Property(x => x.Email).HasMaxLength(200);
            entity.Property(x => x.Direccion).HasMaxLength(300);
            entity.Property(x => x.CodigoPostal).HasMaxLength(10);
            entity.Property(x => x.TipoHuesped).HasMaxLength(30);
            entity.Property(x => x.Empleo).HasMaxLength(100);
            entity.Property(x => x.Situacion).HasMaxLength(100);
            entity.Property(x => x.Finalidad).HasMaxLength(100);
            entity.Property(x => x.EmpleoCategoria).HasMaxLength(100);
        });

        builder.Entity<Reserva>(entity =>
        {
            entity.Property(x => x.Finalidad).HasMaxLength(100);
            entity.Property(x => x.Empleo).HasMaxLength(100);
            entity.Property(x => x.Evaluacion).HasMaxLength(50);
            entity.Property(x => x.PrecioNocheAplicado).HasColumnType("decimal(10,2)");
            entity.Property(x => x.PorcentajeIvaAplicado).HasColumnType("decimal(5,2)");
            entity.Property(x => x.ImporteBase).HasColumnType("decimal(10,2)");
            entity.Property(x => x.ImporteIva).HasColumnType("decimal(10,2)");
            entity.Property(x => x.ImporteTotal).HasColumnType("decimal(10,2)");
            entity.Property(x => x.FormaPago).HasMaxLength(30);
            entity.Property(x => x.Estado).HasConversion<int>();
            entity.HasOne(x => x.Habitacion)
                .WithMany()
                .HasForeignKey(x => x.HabitacionId)
                .OnDelete(DeleteBehavior.Restrict);
            entity.HasOne(x => x.Huesped)
                .WithMany(x => x.Reservas)
                .HasForeignKey(x => x.HuespedId)
                .OnDelete(DeleteBehavior.SetNull);
            entity.HasOne(x => x.Tarifa)
                .WithMany()
                .HasForeignKey(x => x.TarifaId)
                .OnDelete(DeleteBehavior.SetNull);
        });

        builder.Entity<CodigoPostalInfo>(entity =>
        {
            entity.HasKey(x => x.Id);
            entity.Property(x => x.CodigoPostal).IsRequired().HasMaxLength(10);
            entity.Property(x => x.Municipio).IsRequired().HasMaxLength(150);
            entity.Property(x => x.Provincia).IsRequired().HasMaxLength(100);
            entity.HasIndex(x => x.CodigoPostal);
            entity.HasIndex(x => x.Municipio);
        });

        builder.Entity<Factura>(entity =>
        {
            entity.HasIndex(x => new { x.Serie, x.Ejercicio, x.NumeroOrden }).IsUnique();
            entity.Property(x => x.NumeroFactura).IsRequired().HasMaxLength(30);
            entity.Property(x => x.Serie).IsRequired().HasMaxLength(10);
            entity.Property(x => x.BaseImponible).HasColumnType("decimal(10,2)");
            entity.Property(x => x.PorcentajeIva).HasColumnType("decimal(5,2)");
            entity.Property(x => x.CuotaIva).HasColumnType("decimal(10,2)");
            entity.Property(x => x.Total).HasColumnType("decimal(10,2)");
            entity.Property(x => x.Estado).HasConversion<int>();
            entity.HasOne(x => x.Residencia).WithMany().HasForeignKey(x => x.ResidenciaId).OnDelete(DeleteBehavior.Restrict);
            entity.HasOne(x => x.Reserva).WithMany().HasForeignKey(x => x.ReservaId).OnDelete(DeleteBehavior.SetNull);
            entity.HasOne(x => x.Huesped).WithMany().HasForeignKey(x => x.HuespedId).OnDelete(DeleteBehavior.SetNull);
        });

        builder.Entity<LineaFactura>(entity =>
        {
            entity.Property(x => x.PrecioUnidad).HasColumnType("decimal(10,2)");
            entity.Property(x => x.Descuento).HasColumnType("decimal(5,2)");
            entity.Property(x => x.BaseLinea).HasColumnType("decimal(10,2)");
            entity.Property(x => x.PorcentajeIva).HasColumnType("decimal(5,2)");
            entity.Property(x => x.CuotaIvaLinea).HasColumnType("decimal(10,2)");
            entity.Property(x => x.TotalLinea).HasColumnType("decimal(10,2)");
            entity.HasOne(x => x.Factura).WithMany(x => x.Lineas).HasForeignKey(x => x.FacturaId).OnDelete(DeleteBehavior.Cascade);
        });
    }
}
