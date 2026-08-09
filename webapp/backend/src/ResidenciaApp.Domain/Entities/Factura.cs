namespace ResidenciaApp.Domain.Entities;

public enum EstadoFactura
{
    Borrador  = 0,
    Emitida   = 1,
    Pagada    = 2,
    Anulada   = 3
}

public class Factura
{
    public Guid Id { get; set; } = Guid.NewGuid();

    // Numeración (igual que el Excel: GIJ-2026-0001 / SOT-2026-0001 / OVI-2026-0001)
    public string NumeroFactura { get; set; } = string.Empty;   // GIJ-2026-0001
    public string Serie         { get; set; } = string.Empty;   // GIJ / SOT / OVI
    public int    Ejercicio     { get; set; }                    // 2026
    public int    NumeroOrden   { get; set; }                    // 1, 2, 3...

    // Relaciones
    public Guid ResidenciaId { get; set; }
    public Guid? ReservaId   { get; set; }
    public Guid? HuespedId   { get; set; }

    // Fechas
    public DateOnly FechaEmision    { get; set; }
    public DateOnly? FechaVencimiento { get; set; }
    public DateOnly? FechaPago       { get; set; }

    // Snapshot del destinatario (en caso de que cambie)
    public string DestinatarioNombre    { get; set; } = string.Empty;
    public string DestinatarioDni       { get; set; } = string.Empty;
    public string? DestinatarioDireccion { get; set; }
    public string? DestinatarioCp        { get; set; }
    public string? DestinatarioMunicipio { get; set; }

    // Snapshot del emisor (residencia)
    public string EmisorNombre    { get; set; } = string.Empty;
    public string? EmisorCif      { get; set; }
    public string? EmisorDireccion { get; set; }
    public string? EmisorTelefono  { get; set; }

    // Importes
    public decimal BaseImponible    { get; set; } = 0;
    public decimal PorcentajeIva    { get; set; } = 10;
    public decimal CuotaIva         { get; set; } = 0;
    public decimal Total            { get; set; } = 0;

    // Pago
    public string? FormaPago { get; set; }
    public EstadoFactura Estado { get; set; } = EstadoFactura.Borrador;

    // Notas / pie de factura
    public string? Observaciones { get; set; }

    // Auditoría
    public Guid? CreadoPorId { get; set; }
    public DateTime CreadoEn    { get; set; } = DateTime.UtcNow;
    public DateTime ActualizadoEn { get; set; } = DateTime.UtcNow;

    // Navegación
    public Residencia Residencia    { get; set; } = null!;
    public Reserva?   Reserva       { get; set; }
    public Huesped?   Huesped       { get; set; }
    public ICollection<LineaFactura> Lineas { get; set; } = new List<LineaFactura>();
}

public class LineaFactura
{
    public Guid Id        { get; set; } = Guid.NewGuid();
    public Guid FacturaId { get; set; }

    public int     Orden       { get; set; } = 0;
    public string  Concepto    { get; set; } = string.Empty;
    public decimal Cantidad    { get; set; } = 1;
    public string  Unidad      { get; set; } = "noche";        // noche, mes, ud
    public decimal PrecioUnidad { get; set; } = 0;
    public decimal Descuento   { get; set; } = 0;              // % descuento
    public decimal BaseLinea   { get; set; } = 0;
    public decimal PorcentajeIva { get; set; } = 10;
    public decimal CuotaIvaLinea { get; set; } = 0;
    public decimal TotalLinea  { get; set; } = 0;

    public Factura Factura { get; set; } = null!;
}
