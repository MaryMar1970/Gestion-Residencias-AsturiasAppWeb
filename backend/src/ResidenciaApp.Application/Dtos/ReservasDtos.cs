using System.ComponentModel.DataAnnotations;

namespace ResidenciaApp.Application.Dtos;

// ─── HUÉSPED ──────────────────────────────────────────────────────────────────
public record HuespedDto(
    Guid Id, string Dni, string Nombre, string Apellidos,
    string NombreCompleto,
    string? Telefono, string? Email,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? CentroOrigen, string? Departamento,
    string TipoHuesped, bool EnListaNegra, string? MotivoListaNegra,
    string? Notas, string? Empleo, string? Situacion, string? Finalidad, string? EmpleoCategoria,
    string FamiliaNumerosa, decimal PorcentajeDescuento,
    int TotalReservas, DateTime CreadoEn);

public record UpsertHuespedDto(
    [Required(ErrorMessage = "El DNI es obligatorio"), StringLength(20)] string Dni,
    [Required(ErrorMessage = "El Nombre es obligatorio"), StringLength(100)] string Nombre,
    [Required(ErrorMessage = "Los Apellidos son obligatorios"), StringLength(100)] string Apellidos,
    [StringLength(50)] string? Telefono,
    [EmailAddress, StringLength(150)] string? Email,
    [StringLength(200)] string? Direccion,
    [StringLength(10)] string? CodigoPostal,
    [StringLength(100)] string? Municipio,
    [StringLength(100)] string? Provincia,
    [StringLength(100)] string? CentroOrigen,
    [StringLength(100)] string? Departamento,
    [StringLength(50)] string? TipoHuesped,
    bool EnListaNegra,
    string? MotivoListaNegra,
    string? Notas,
    string? Empleo,
    string? Situacion,
    string? Finalidad,
    string? EmpleoCategoria,
    string? FamiliaNumerosa,
    [Range(0, 100)] decimal? PorcentajeDescuento);

// ─── RESERVA ──────────────────────────────────────────────────────────────────
public record ReservaDto(
    Guid Id, int NumeroOrden, Guid? ResidenciaId,
    Guid? HabitacionId, string? HabitacionNumero, string? ResidenciaNombre, string? TipoHabitacionNombre,
    Guid? HuespedId, string? HuespedNombreCompleto, string? HuespedDni, string? HuespedNombre, string? HuespedApellidos,
    string? HuespedTelefono, string? HuespedEmail,
    string FechaEntrada, string FechaSalida,
    int TotalNoches, int NumPersonas, int NumNinos, int CamasSupletorias,
    string FamiliaNumerosa, decimal PorcentajeDescuento,
    string? AlojamientoSolicitado,
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
    Guid? ResidenciaId,
    Guid? HabitacionId,
    Guid? HuespedId,
    int? NumeroOrden,
    [Required(ErrorMessage = "La fecha de entrada es obligatoria")] string FechaEntrada,   // "yyyy-MM-dd"
    [Required(ErrorMessage = "La fecha de salida es obligatoria")] string FechaSalida,    // "yyyy-MM-dd"
    [Range(1, 20, ErrorMessage = "El número de personas debe estar entre 1 y 20")] int NumPersonas,
    [Range(0, 10)] int NumNinos,
    [Range(0, 5)] int CamasSupletorias,
    string? FamiliaNumerosa,
    [Range(0, 100)] decimal? PorcentajeDescuento,
    string? AlojamientoSolicitado,
    bool EsBloqueo,
    string? MotivoBloqueo,
    Guid? TarifaId,
    string? Observaciones,
    string? Finalidad,
    string? Resolucion,
    DateTime? FechaSolicitud);


public record ActualizarReservaDto(
    Guid? ResidenciaId,
    [Required(ErrorMessage = "La fecha de entrada es obligatoria")] string FechaEntrada,
    [Required(ErrorMessage = "La fecha de salida es obligatoria")] string FechaSalida,
    [Range(1, 20)] int NumPersonas,
    [Range(0, 10)] int NumNinos,
    [Range(0, 5)] int CamasSupletorias,
    string? FamiliaNumerosa,
    [Range(0, 100)] decimal? PorcentajeDescuento,
    string? AlojamientoSolicitado,
    [Required(ErrorMessage = "El estado es obligatorio")] string Estado,
    bool EsBloqueo,
    string? MotivoBloqueo,
    Guid? TarifaId,
    string? Observaciones,
    bool Pagado,
    string? FormaPago,
    string? Finalidad,
    string? Empleo,
    string? Situacion,
    Guid? HabitacionId,
    string? Resolucion,
    DateTime? FechaSolicitud,
    Guid? ReevaluarCandidatoId);



// ─── CALENDARIO ───────────────────────────────────────────────────────────────
public record CalendarioHabitacionDto(
    Guid HabitacionId,
    Guid ResidenciaId,
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

// ─── RESULTADO OPERACIÓN RESERVA (CON REEVALUADOS) ───────────────────────────
public record ResultadoOperacionReservaDto(
    ReservaDto Reserva,
    IList<ReservaDto> ReservasReevaluadas);

public record MoverReservaDto(
    Guid HabitacionId,
    string FechaEntrada,
    string FechaSalida);



