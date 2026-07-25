namespace ResidenciaApp.Domain.Entities;

public class Habitacion
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public Guid ResidenciaId { get; set; }
    public Guid TipoHabitacionId { get; set; }
    public string Numero { get; set; } = string.Empty;       // "1", "OF.1", "Ap.2", "EST.1"
    public string? Nombre { get; set; }                       // Nombre descriptivo opcional
    public int CapacidadPersonas { get; set; } = 1;          // Puede sobreescribir el del tipo
    public bool AdmiteSupletorias { get; set; } = false;
    public int PlazasSupletorias { get; set; } = 0;
    public bool Activa { get; set; } = true;                  // false = oculta/bloqueada permanentemente
    public string? Notas { get; set; }
    public int Orden { get; set; } = 0;                       // Orden en el calendario
    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
    public DateTime ActualizadoEn { get; set; } = DateTime.UtcNow;

    public Residencia Residencia { get; set; } = null!;
    public TipoHabitacion TipoHabitacion { get; set; } = null!;
}
