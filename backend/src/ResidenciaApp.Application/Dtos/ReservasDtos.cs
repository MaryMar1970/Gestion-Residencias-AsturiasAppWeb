namespace ResidenciaApp.Application.Dtos;

// ─── HUÉSPED ──────────────────────────────────────────────────────────────────
public record HuespedDto(
    Guid Id, string Dni, string Nombre, string Apellidos,
    string NombreCompleto,
    string? Telefono, string? Email,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? CentroOrigen, string? Departamento,
    string TipoHuesped, bool EnListaNegra, string? MotivoListaNegra,
    string? Notas, string? Empleo, string? Situacion, string? Finalidad, string? EmpleoCategoria, int TotalReservas, DateTime CreadoEn);

public record UpsertHuespedDto(
    string Dni, string Nombre, string Apellidos,
    string? Telefono, string? Email,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? CentroOrigen, string? Departamento,
    string TipoHuesped, bool EnListaNegra, string? MotivoListaNegra,
    string? Notas, string? Empleo, string? Situacion, string? Finalidad, string? EmpleoCategoria);

// ─── RESERVA ──────────────────────────────────────────────────────────────────
public record ReservaDto(
    Guid Id, int NumeroOrden,
    Guid HabitacionId, string HabitacionNumero, string ResidenciaNombre, string TipoHabitacionNombre,
    Guid? HuespedId, string? HuespedNombreCompleto, string? HuespedDni, string? HuespedNombre, string? HuespedApellidos,
    string? HuespedTelefono, string? HuespedEmail,
    string FechaEntrada, string FechaSalida,
    int TotalNoches, int NumPersonas, int CamasSupletorias,
    string Estado, int EstadoInt,
    bool EsBloqueo, string? MotivoBloqueo,
    string? TarifaNombreSnapshot,
    decimal PrecioNocheAplicado, decimal PorcentajeIvaAplicado,
    decimal ImporteBase, decimal ImporteIva, decimal ImporteTotal,
    bool Pagado, string? FormaPago, DateTime? FechaPago,
    bool Facturado, string? Observaciones,
    string? Finalidad, string? Empleo, string? Evaluacion,
    string Resolucion, DateTime? FechaSolicitud,
    DateTime CreadoEn, DateTime ActualizadoEn);

public record CrearReservaDto(
    Guid? HabitacionId,
    Guid? HuespedId,
    string FechaEntrada,   // "yyyy-MM-dd"
    string FechaSalida,    // "yyyy-MM-dd"
    int NumPersonas,
    int CamasSupletorias,
    bool EsBloqueo,
    string? MotivoBloqueo,
    Guid? TarifaId,
    string? Observaciones,
    string? Finalidad,
    string? Resolucion,
    DateTime? FechaSolicitud);

public record ActualizarReservaDto(
    string FechaEntrada,
    string FechaSalida,
    int NumPersonas,
    int CamasSupletorias,
    string Estado,
    bool EsBloqueo,
    string? MotivoBloqueo,
    Guid? TarifaId,
    string? Observaciones,
    bool Pagado,
    string? FormaPago,
    string? Finalidad,
    string? Resolucion,
    DateTime? FechaSolicitud);

// ─── CALENDARIO ───────────────────────────────────────────────────────────────
public record CalendarioHabitacionDto(
    Guid HabitacionId,
    string Numero,
    string ResidenciaNombre,
    string TipoNombre,
    string TipoCodigo,
    bool Activa,
    int Orden,
    IList<ReservaDto> Reservas);

public record CalendarioDto(
    string FechaInicio,
    string FechaFin,
    IList<CalendarioHabitacionDto> Habitaciones);

// ─── SOLAPAMIENTO ─────────────────────────────────────────────────────────────
public record SolapamientoDto(
    bool HaySolapamiento,
    string? Mensaje,
    IList<ReservaDto> ReservasConflictivas);
