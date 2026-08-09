namespace ResidenciaApp.Domain.Entities;

public class Festivo
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public DateOnly Fecha { get; set; }
    public string Descripcion { get; set; } = string.Empty;
    public string Ambito { get; set; } = "NACIONAL";   // NACIONAL, REGIONAL, LOCAL
    public bool Activo { get; set; } = true;
    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
}
