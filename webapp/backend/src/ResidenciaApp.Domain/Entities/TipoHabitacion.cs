namespace ResidenciaApp.Domain.Entities;

public class TipoHabitacion
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string Nombre { get; set; } = string.Empty;       // "Individual", "Doble", "Apartamento"...
    public string Codigo { get; set; } = string.Empty;       // "IND", "DOB", "APT", "TRI", "CUA"
    public int CapacidadMaxima { get; set; } = 1;            // Personas máximas
    public bool AdmiteSupletorias { get; set; } = false;
    public string? Descripcion { get; set; }
    public bool Activo { get; set; } = true;
    public int Orden { get; set; } = 0;
    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;

    public ICollection<Habitacion> Habitaciones { get; set; } = new List<Habitacion>();
    public ICollection<Tarifa> Tarifas { get; set; } = new List<Tarifa>();
}
