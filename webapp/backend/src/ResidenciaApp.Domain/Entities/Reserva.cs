namespace ResidenciaApp.Domain.Entities;

public enum EstadoReserva
{
    Pendiente = 0,
    Confirmada = 1,
    CheckIn = 2,
    CheckOut = 3,
    Cancelada = 4,
    Bloqueada = 5   // habitación bloqueada, sin huésped
}

public static class Resoluciones
{
    public const string Si = "SI";
    public const string No = "NO";
    public const string Concedida = "CONCEDIDA";
    public const string Denegada = "DENEGADA";
    public const string Desestimada = "DESESTIMADA";
    public const string Renuncia = "RENUNCIA";
    public const string Reevaluada = "REEVALUADA";

    public static bool EsNegativa(string? r) =>
        r == No || r == Denegada || r == Desestimada || r == Renuncia;

    public static bool EsPositivaOValida(string? r) =>
        string.IsNullOrEmpty(r) || r == Si || r == Concedida || r == Reevaluada;
}

public class Reserva
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public int NumeroOrden { get; set; }           // Número de orden secuencial autoincremental (Nº ORDEN)

    // Relaciones
    public Guid? ResidenciaId { get; set; }        // Residencia a la que pertenece la solicitud
    public Residencia? Residencia { get; set; }
    public Guid? HabitacionId { get; set; }        // null si la solicitud no tiene habitación asignada todavía o fue denegada
    public Guid? HuespedId { get; set; }           // null si es bloqueo sin huésped

    // Fechas
    public DateOnly FechaEntrada { get; set; }
    public DateOnly FechaSalida { get; set; }

    // Estado y Prioridad
    public EstadoReserva Estado { get; set; } = EstadoReserva.Pendiente;
    public string? Finalidad { get; set; }         // Comisión, Destino, Otros, etc.
    public string? Empleo { get; set; }            // Categoría de empleo mapeada (GC, Alumno, etc.)
    public string? Evaluacion { get; set; }        // Clasificación de prioridad calculada (ej. 1, 1, 2)

    // Ocupación
    public int NumPersonas { get; set; } = 1;
    public int NumNinos { get; set; } = 0;
    public int CamasSupletorias { get; set; } = 0;
    public string FamiliaNumerosa { get; set; } = "NO";      // "NO", "GENERAL", "ESPECIAL"
    public decimal PorcentajeDescuento { get; set; } = 0;    // 0, 20, 50%
    public string? AlojamientoSolicitado { get; set; }       // Petición preferencia de alojamiento del huésped

    // Tipo de reserva / bloqueo
    public bool EsBloqueo { get; set; } = false;
    public string? MotivoBloqueo { get; set; }

    // Tarifa aplicada
    public Guid? TarifaId { get; set; }
    public string? TarifaNombreSnapshot { get; set; }  // copia por si cambia la tarifa
    public decimal PrecioNocheAplicado { get; set; } = 0;
    public decimal PorcentajeIvaAplicado { get; set; } = 10;

    // Cálculos (se calculan al crear/modificar)
    public int TotalNoches { get; set; } = 0;
    public decimal ImporteBase { get; set; } = 0;
    public decimal ImporteIva { get; set; } = 0;
    public decimal ImporteTotal { get; set; } = 0;

    // Facturación
    public bool Facturado { get; set; } = false;
    public Guid? FacturaId { get; set; }

    // Pago
    public bool Pagado { get; set; } = false;
    public string? FormaPago { get; set; }         // Efectivo, Tarjeta, Bizum, Transferencia
    public DateTime? FechaPago { get; set; }

    // Observaciones
    public string? Observaciones { get; set; }
    public string Resolucion { get; set; } = "SI";
    public DateTime? FechaSolicitud { get; set; }

    // Auditoría
    public Guid? CreadoPorId { get; set; }
    public DateTime CreadoEn { get; set; } = DateTime.UtcNow;
    public DateTime ActualizadoEn { get; set; } = DateTime.UtcNow;

    // Navegación
    public Habitacion? Habitacion { get; set; }
    public Huesped? Huesped { get; set; }
    public Tarifa? Tarifa { get; set; }
}
