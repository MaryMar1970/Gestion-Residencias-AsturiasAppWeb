namespace ResidenciaApp.Application.Dtos;

// ─── RESIDENCIA ────────────────────────────────────────────────────────────────
public record ResidenciaDto(
    Guid Id, string Nombre, string? RazonSocial, string? Cif,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? Telefono, string? Email, bool Activa, int Orden,
    int TotalHabitaciones);

public record UpsertResidenciaDto(
    string Nombre, string? RazonSocial, string? Cif,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? Telefono, string? Email, bool Activa, int Orden);

// ─── TIPO HABITACIÓN ──────────────────────────────────────────────────────────
public record TipoHabitacionDto(
    Guid Id, string Nombre, string Codigo, int CapacidadMaxima,
    bool AdmiteSupletorias, string? Descripcion, bool Activo, int Orden);

public record UpsertTipoHabitacionDto(
    string Nombre, string Codigo, int CapacidadMaxima,
    bool AdmiteSupletorias, string? Descripcion, bool Activo, int Orden);

// ─── HABITACIÓN ────────────────────────────────────────────────────────────────
public record HabitacionDto(
    Guid Id, Guid ResidenciaId, string ResidenciaNombre,
    Guid TipoHabitacionId, string TipoNombre, string TipoCodigo,
    string Numero, string? Nombre, int CapacidadPersonas,
    bool AdmiteSupletorias, int PlazasSupletorias,
    bool Activa, string? Notas, int Orden);

public record UpsertHabitacionDto(
    Guid ResidenciaId, Guid TipoHabitacionId,
    string Numero, string? Nombre, int CapacidadPersonas,
    bool AdmiteSupletorias, int PlazasSupletorias,
    bool Activa, string? Notas, int Orden);

// ─── TARIFA ───────────────────────────────────────────────────────────────────
public record TarifaDto(
    Guid Id, Guid ResidenciaId, string ResidenciaNombre, Guid TipoHabitacionId, string TipoHabitacionNombre,
    string NombreTarifa, decimal PrecioNoche, decimal PrecioMes,
    decimal PorcentajeIva, string? Descripcion, bool Activa,
    DateTime? VigenteDesde, DateTime? VigenteHasta);

public record UpsertTarifaDto(
    Guid ResidenciaId, Guid TipoHabitacionId, string NombreTarifa,
    decimal PrecioNoche, decimal PrecioMes, decimal PorcentajeIva,
    string? Descripcion, bool Activa,
    DateTime? VigenteDesde, DateTime? VigenteHasta);

// ─── FESTIVO ──────────────────────────────────────────────────────────────────
public record FestivoDto(
    Guid Id, DateOnly Fecha, string Descripcion, string Ambito, bool Activo);

public record UpsertFestivoDto(
    DateOnly Fecha, string Descripcion, string Ambito, bool Activo);

// ─── USUARIO (para gestión) ───────────────────────────────────────────────────
public record UsuarioListaDto(
    Guid Id, string Email, string FullName,
    IList<string> Roles, bool Activo, DateTime CreadoEn);

public record CrearUsuarioDto(
    string Email, string FullName, string Password, string Rol);

public record ActualizarUsuarioDto(
    string FullName, string? NuevaPassword, string Rol, bool Activo);
