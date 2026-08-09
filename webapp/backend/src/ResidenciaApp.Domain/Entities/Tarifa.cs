namespace ResidenciaApp.Domain.Entities;

public class Tarifa
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid TipoHabitacionId { get; set; }
    public string NombreTarifa { get; set; } = string.Empty;       // "Estudiante", "Investigador", "Externo"
    public decimal PrecioNoche { get; set; } = 0;
    public decimal PrecioMes { get; set; } = 0;
    public Guid ResidenciaId { get; set; }
    public decimal PorcentajeIva { get; set; } = 10;               // 10% por defecto (alojamiento)
    public string? Descripcion { get; set; }
    public bool Activa { get; set; } = true;
    public DateTime? VigenteDesde { get; set; }
    public DateTime? VigenteHasta { get; set; }
    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
    public DateTime ActualizadoEn { get; set; } = DateTime.UtcNow;

    public TipoHabitacion TipoHabitacion { get; set; } = null!;
    public Residencia Residencia { get; set; } = null!;
}
