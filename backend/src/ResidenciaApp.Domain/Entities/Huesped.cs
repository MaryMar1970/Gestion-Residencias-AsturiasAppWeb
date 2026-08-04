namespace ResidenciaApp.Domain.Entities;

public class Huesped
{
    public Guid Id { get; set; } = Guid.NewGuid();

    // Identificación
    public string Dni { get; set; } = string.Empty;
    public string Nombre { get; set; } = string.Empty;
    public string Apellidos { get; set; } = string.Empty;

    // Contacto
    public string? Telefono { get; set; }
    public string? Email { get; set; }

    // Dirección
    public string? Direccion { get; set; }
    public string? CodigoPostal { get; set; }
    public string? Municipio { get; set; }
    public string? Provincia { get; set; }

    // Centro/Entidad de origen (campo del Excel)
    public string? CentroOrigen { get; set; }
    public string? Departamento { get; set; }

    // Tipo de huésped e información militar/laboral
    public string TipoHuesped { get; set; } = "Externo"; // Estudiante, Investigador, Externo, Trabajador
    public string? Empleo { get; set; }                      // Guardia, Cabo, Sargento, etc. (Rango)
    public string? Situacion { get; set; }                   // Activo, Viogen, Reserva, etc.
    public string? Finalidad { get; set; }                   // Otros, Comisión, etc. (Default finality)
    public string? EmpleoCategoria { get; set; }             // GC, Alumno, etc. (Default category)

    // Descuentos y Bonificaciones
    public string FamiliaNumerosa { get; set; } = "NO";      // "NO", "GENERAL" (20%), "ESPECIAL" (50%)
    public decimal PorcentajeDescuento { get; set; } = 0;    // 0, 20, 50

    // Lista negra
    public bool EnListaNegra { get; set; } = false;
    public string? MotivoListaNegra { get; set; }

    // Notas internas
    public string? Notas { get; set; }

    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
    public DateTime ActualizadoEn { get; set; } = DateTime.UtcNow;

    public ICollection<Reserva> Reservas { get; set; } = new List<Reserva>();
}
