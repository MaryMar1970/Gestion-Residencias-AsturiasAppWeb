namespace ResidenciaApp.Application.Dtos;

// ─── LÍNEAS ────────────────────────────────────────────────────────────────────
public record LineaFacturaDto(
    Guid Id, int Orden, string Concepto,
    decimal Cantidad, string Unidad, decimal PrecioUnidad, decimal Descuento,
    decimal BaseLinea, decimal PorcentajeIva, decimal CuotaIvaLinea, decimal TotalLinea);

public record UpsertLineaFacturaDto(
    int Orden, string Concepto,
    decimal Cantidad, string Unidad, decimal PrecioUnidad, decimal Descuento,
    decimal PorcentajeIva);

// ─── FACTURA ───────────────────────────────────────────────────────────────────
public record FacturaDto(
    Guid Id,
    string NumeroFactura, string Serie, int Ejercicio, int NumeroOrden,
    Guid ResidenciaId, string ResidenciaNombre,
    Guid? ReservaId, Guid? HuespedId,
    string FechaEmision, string? FechaVencimiento, string? FechaPago,
    string DestinatarioNombre, string DestinatarioDni,
    string? DestinatarioDireccion, string? DestinatarioCp, string? DestinatarioMunicipio,
    string EmisorNombre, string? EmisorCif, string? EmisorDireccion, string? EmisorTelefono,
    decimal BaseImponible, decimal PorcentajeIva, decimal CuotaIva, decimal Total,
    string? FormaPago, string Estado, int EstadoInt,
    string? Observaciones,
    IList<LineaFacturaDto> Lineas,
    DateTime CreadoEn);

public record CrearFacturaDto(
    Guid ResidenciaId,
    Guid? ReservaId,          // si viene de una reserva → auto-rellena lineas
    Guid? HuespedId,
    string FechaEmision,      // "yyyy-MM-dd"
    string? FechaVencimiento,
    // Destinatario (se puede sobreescribir el auto-relleno del huésped)
    string DestinatarioNombre,
    string DestinatarioDni,
    string? DestinatarioDireccion,
    string? DestinatarioCp,
    string? DestinatarioMunicipio,
    decimal PorcentajeIva,
    string? FormaPago,
    string? Observaciones,
    IList<UpsertLineaFacturaDto> Lineas);

public record EmitirFacturaDto(string? FormaPago);

public record AnularFacturaDto(string? Motivo);
