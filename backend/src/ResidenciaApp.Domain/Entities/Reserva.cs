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

public class Reserva
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public int NumeroOrden { get; set; }           // Número de orden secuencial autoincremental (Nº ORDEN)

    // Relaciones
    public Guid HabitacionId { get; set; }
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
    public int CamasSupletorias { get; set; } = 0;

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
    public Habitacion Habitacion { get; set; } = null!;
    public Huesped? Huesped { get; set; }
    public Tarifa? Tarifa { get; set; }
}
