using System.ComponentModel.DataAnnotations;

namespace ResidenciaApp.Application.Dtos;

// ─── RESIDENCIA ────────────────────────────────────────────────────────────────
public record ResidenciaDto(
    Guid Id, string Nombre, string? RazonSocial, string? Cif,
    string? Direccion, string? CodigoPostal, string? Municipio, string? Provincia,
    string? Telefono, string? Email, string? SerieFactura, bool Activa, int Orden,
    int TotalHabitaciones);

public record UpsertResidenciaDto(
    [Required(ErrorMessage = "El nombre de la residencia es obligatorio"), StringLength(150)] string Nombre,
    [StringLength(200)] string? RazonSocial,
    [StringLength(20)] string? Cif,
    [StringLength(200)] string? Direccion,
    [StringLength(10)] string? CodigoPostal,
    [StringLength(100)] string? Municipio,
    [StringLength(100)] string? Provincia,
    [StringLength(50)] string? Telefono,
    [EmailAddress, StringLength(150)] string? Email,
    [StringLength(10)] string? SerieFactura,
    bool Activa,
    int Orden);

// ─── TIPO HABITACIÓN ──────────────────────────────────────────────────────────
public record TipoHabitacionDto(
    Guid Id, string Nombre, string Codigo, int CapacidadMaxima,
    bool AdmiteSupletorias, string? Descripcion, bool Activo, int Orden);

public record UpsertTipoHabitacionDto(
    [Required(ErrorMessage = "El nombre es obligatorio"), StringLength(100)] string Nombre,
    [Required(ErrorMessage = "El código es obligatorio"), StringLength(20)] string Codigo,
    [Range(1, 20)] int CapacidadMaxima,
    bool AdmiteSupletorias,
    string? Descripcion,
    bool Activo,
    int Orden);

// ─── HABITACIÓN ────────────────────────────────────────────────────────────────
public record HabitacionDto(
    Guid Id, Guid ResidenciaId, string ResidenciaNombre,
    Guid TipoHabitacionId, string TipoNombre, string TipoCodigo,
    string Numero, string? Nombre, string TipoCamaPrincipal, int CapacidadPersonas,
    bool AdmiteSupletorias, int PlazasSupletorias,
    bool Activa, string? Notas, int Orden);

public record UpsertHabitacionDto(
    [Required] Guid ResidenciaId,
    [Required] Guid TipoHabitacionId,
    [Required(ErrorMessage = "El número de habitación es obligatorio"), StringLength(20)] string Numero,
    [StringLength(100)] string? Nombre,
    string TipoCamaPrincipal,
    [Range(1, 20)] int CapacidadPersonas,
    bool AdmiteSupletorias,
    [Range(0, 10)] int PlazasSupletorias,
    bool Activa,
    string? Notas,
    int Orden);

public record CrearLoteHabitacionesDto(
    [Required] Guid ResidenciaId,
    [Required] Guid TipoHabitacionId,
    [Required, StringLength(10)] string Prefijo,
    [Range(1, 9999)] int NumeroInicio,
    [Range(1, 100)] int Cantidad,
    string TipoCamaPrincipal,
    [Range(1, 20)] int CapacidadPersonas,
    bool AdmiteSupletorias,
    [Range(0, 10)] int PlazasSupletorias);

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
