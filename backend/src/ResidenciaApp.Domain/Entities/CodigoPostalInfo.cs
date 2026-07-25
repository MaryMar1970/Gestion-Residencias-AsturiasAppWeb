namespace ResidenciaApp.Domain.Entities;

public class CodigoPostalInfo
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string CodigoPostal { get; set; } = string.Empty;
    public string Municipio { get; set; } = string.Empty;
    public string Provincia { get; set; } = string.Empty;
}
