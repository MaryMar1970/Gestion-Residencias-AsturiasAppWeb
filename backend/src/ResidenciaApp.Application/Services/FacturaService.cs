using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public interface IFacturaService
{
    Task<List<FacturaDto>> GetAllAsync(Guid? residenciaId = null, int? ejercicio = null, string? estado = null);
    Task<FacturaDto?> GetByIdAsync(Guid id);
    Task<FacturaDto?> GetByNumeroAsync(string numero);
    Task<FacturaDto> CrearDesdeReservaAsync(Guid reservaId, string? formaPago = null);
    Task<FacturaDto> CrearManualAsync(CrearFacturaDto dto);
    Task<FacturaDto?> EmitirAsync(Guid id, EmitirFacturaDto dto);
    Task<FacturaDto?> MarcarPagadaAsync(Guid id, string? formaPago = null);
    Task<FacturaDto?> AnularAsync(Guid id, AnularFacturaDto dto);
    Task<string> GenerarHtmlAsync(Guid id);   // HTML para imprimir / PDF
}

public class FacturaService(IResidenciaDbContext db) : IFacturaService
{
    // ─── helpers ──────────────────────────────────────────────────────────────
    private static LineaFacturaDto ToLineaDto(LineaFactura l) => new(
        l.Id, l.Orden, l.Concepto, l.Cantidad, l.Unidad,
        l.PrecioUnidad, l.Descuento, l.BaseLinea,
        l.PorcentajeIva, l.CuotaIvaLinea, l.TotalLinea);

    private static FacturaDto ToDto(Factura f) => new(
        f.Id, f.NumeroFactura, f.Serie, f.Ejercicio, f.NumeroOrden,
        f.ResidenciaId, f.Residencia?.Nombre ?? "",
        f.ReservaId, f.HuespedId,
        f.FechaEmision.ToString("yyyy-MM-dd"),
        f.FechaVencimiento?.ToString("yyyy-MM-dd"),
        f.FechaPago?.ToString("yyyy-MM-dd"),
        f.DestinatarioNombre, f.DestinatarioDni,
        f.DestinatarioDireccion, f.DestinatarioCp, f.DestinatarioMunicipio,
        f.EmisorNombre, f.EmisorCif, f.EmisorDireccion, f.EmisorTelefono,
        f.BaseImponible, f.PorcentajeIva, f.CuotaIva, f.Total,
        f.FormaPago, f.Estado.ToString(), (int)f.Estado,
        f.Observaciones,
        f.Lineas.OrderBy(l => l.Orden).Select(ToLineaDto).ToList(),
        f.CreadoEn);

    private static IQueryable<Factura> WithIncludes(IQueryable<Factura> q) =>
        q.Include(f => f.Residencia)
         .Include(f => f.Huesped)
         .Include(f => f.Reserva)
         .Include(f => f.Lineas);

    // Genera el número de factura: GIJ-2026-0001
    private async Task<(string serie, int orden, string numero)> SiguienteNumeroAsync(Guid residenciaId, int ejercicio)
    {
        var res = await db.Residencias.FindAsync(residenciaId)
            ?? throw new InvalidOperationException("Residencia no encontrada.");

        // Serie = primeras 3 letras del nombre sin acentos, en mayúsculas
        var serie = new string(res.Nombre.Normalize(System.Text.NormalizationForm.FormD)
            .Where(c => c < 128 && char.IsLetter(c))
            .Take(3)
            .ToArray()).ToUpper();

        var ultimoOrden = await db.Facturas
            .Where(f => f.Serie == serie && f.Ejercicio == ejercicio)
            .MaxAsync(f => (int?)f.NumeroOrden) ?? 0;

        var orden = ultimoOrden + 1;
        var numero = $"{serie}-{ejercicio}-{orden:D4}";
        return (serie, orden, numero);
    }

    // Calcula totales de las líneas
    private static void CalcularLinea(LineaFactura l)
    {
        var bruto = l.Cantidad * l.PrecioUnidad;
        l.BaseLinea    = Math.Round(bruto * (1 - l.Descuento / 100), 2);
        l.CuotaIvaLinea = Math.Round(l.BaseLinea * (l.PorcentajeIva / 100), 2);
        l.TotalLinea   = l.BaseLinea + l.CuotaIvaLinea;
    }

    private static void RecalcularTotales(Factura f)
    {
        f.BaseImponible = Math.Round(f.Lineas.Sum(l => l.BaseLinea), 2);
        f.CuotaIva      = Math.Round(f.Lineas.Sum(l => l.CuotaIvaLinea), 2);
        f.Total         = f.BaseImponible + f.CuotaIva;
        // IVA dominante (del primer grupo)
        f.PorcentajeIva = f.Lineas.FirstOrDefault()?.PorcentajeIva ?? 10;
    }

    // ─── CRUD ─────────────────────────────────────────────────────────────────
    public async Task<List<FacturaDto>> GetAllAsync(Guid? residenciaId = null, int? ejercicio = null, string? estado = null)
    {
        var q = WithIncludes(db.Facturas.AsQueryable());
        if (residenciaId.HasValue) q = q.Where(f => f.ResidenciaId == residenciaId.Value);
        if (ejercicio.HasValue)    q = q.Where(f => f.Ejercicio == ejercicio.Value);
        if (!string.IsNullOrEmpty(estado) && Enum.TryParse<EstadoFactura>(estado, out var est))
            q = q.Where(f => f.Estado == est);
        return await q.OrderByDescending(f => f.FechaEmision).ThenByDescending(f => f.NumeroOrden)
            .Select(f => ToDto(f)).ToListAsync();
    }

    public async Task<FacturaDto?> GetByIdAsync(Guid id)
    {
        var f = await WithIncludes(db.Facturas.Where(x => x.Id == id)).FirstOrDefaultAsync();
        return f is null ? null : ToDto(f);
    }

    public async Task<FacturaDto?> GetByNumeroAsync(string numero)
    {
        var f = await WithIncludes(db.Facturas.Where(x => x.NumeroFactura == numero)).FirstOrDefaultAsync();
        return f is null ? null : ToDto(f);
    }

    // ─── CREAR DESDE RESERVA ──────────────────────────────────────────────────
    public async Task<FacturaDto> CrearDesdeReservaAsync(Guid reservaId, string? formaPago = null)
    {
        var reserva = await db.Reservas
            .Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
            .Include(r => r.Huesped)
            .Include(r => r.Tarifa)
            .FirstOrDefaultAsync(r => r.Id == reservaId)
            ?? throw new InvalidOperationException("Reserva no encontrada.");

        if (reserva.Facturado)
            throw new InvalidOperationException("Esta reserva ya tiene una factura emitida.");

        var residencia = reserva.Habitacion.Residencia;
        var huesped    = reserva.Huesped;
        var hoy        = DateOnly.FromDateTime(DateTime.Today);
        var ejercicio  = hoy.Year;

        var (serie, orden, numero) = await SiguienteNumeroAsync(residencia.Id, ejercicio);

        // Línea principal: alojamiento
        var lineaAloj = new LineaFactura
        {
            Orden       = 1,
            Concepto    = $"Alojamiento {reserva.Habitacion.Numero} — {reserva.FechaEntrada:dd/MM/yyyy} a {reserva.FechaSalida:dd/MM/yyyy}",
            Cantidad    = reserva.TotalNoches,
            Unidad      = "noche",
            PrecioUnidad = reserva.PrecioNocheAplicado,
            Descuento   = 0,
            PorcentajeIva = reserva.PorcentajeIvaAplicado
        };
        CalcularLinea(lineaAloj);

        var factura = new Factura
        {
            NumeroFactura = numero, Serie = serie, Ejercicio = ejercicio, NumeroOrden = orden,
            ResidenciaId  = residencia.Id,
            ReservaId     = reserva.Id,
            HuespedId     = huesped?.Id,
            FechaEmision  = hoy,
            FechaVencimiento = hoy.AddDays(30),
            DestinatarioNombre    = huesped is null ? "—" : $"{huesped.Nombre} {huesped.Apellidos}",
            DestinatarioDni       = huesped?.Dni ?? "—",
            DestinatarioDireccion = null,
            DestinatarioCp        = huesped?.CodigoPostal,
            DestinatarioMunicipio = huesped?.Municipio,
            EmisorNombre    = residencia.Nombre,
            EmisorCif       = residencia.Cif,
            EmisorDireccion = residencia.Direccion,
            EmisorTelefono  = residencia.Telefono,
            FormaPago       = formaPago ?? reserva.FormaPago,
            Estado          = EstadoFactura.Emitida,
            Lineas          = new List<LineaFactura> { lineaAloj }
        };
        RecalcularTotales(factura);

        db.Facturas.Add(factura);

        // Marcar reserva como facturada
        reserva.Facturado = true;
        reserva.FacturaId = factura.Id;
        reserva.ActualizadoEn = DateTime.UtcNow;

        await db.SaveChangesAsync();
        return (await GetByIdAsync(factura.Id))!;
    }

    // ─── CREAR MANUAL ─────────────────────────────────────────────────────────
    public async Task<FacturaDto> CrearManualAsync(CrearFacturaDto dto)
    {
        var hoy = DateOnly.Parse(dto.FechaEmision);
        var (serie, orden, numero) = await SiguienteNumeroAsync(dto.ResidenciaId, hoy.Year);

        var residencia = await db.Residencias.FindAsync(dto.ResidenciaId)!;

        var lineas = dto.Lineas.Select((l, i) =>
        {
            var linea = new LineaFactura
            {
                Orden = l.Orden > 0 ? l.Orden : i + 1,
                Concepto = l.Concepto, Cantidad = l.Cantidad, Unidad = l.Unidad,
                PrecioUnidad = l.PrecioUnidad, Descuento = l.Descuento, PorcentajeIva = l.PorcentajeIva
            };
            CalcularLinea(linea);
            return linea;
        }).ToList();

        var factura = new Factura
        {
            NumeroFactura = numero, Serie = serie, Ejercicio = hoy.Year, NumeroOrden = orden,
            ResidenciaId  = dto.ResidenciaId,
            ReservaId     = dto.ReservaId,
            HuespedId     = dto.HuespedId,
            FechaEmision  = hoy,
            FechaVencimiento = dto.FechaVencimiento is null ? null : DateOnly.Parse(dto.FechaVencimiento),
            DestinatarioNombre    = dto.DestinatarioNombre,
            DestinatarioDni       = dto.DestinatarioDni,
            DestinatarioDireccion = dto.DestinatarioDireccion,
            DestinatarioCp        = dto.DestinatarioCp,
            DestinatarioMunicipio = dto.DestinatarioMunicipio,
            EmisorNombre    = residencia!.Nombre,
            EmisorCif       = residencia.Cif,
            EmisorDireccion = residencia.Direccion,
            EmisorTelefono  = residencia.Telefono,
            FormaPago       = dto.FormaPago,
            Observaciones   = dto.Observaciones,
            Estado          = EstadoFactura.Borrador,
            Lineas          = lineas
        };
        RecalcularTotales(factura);
        db.Facturas.Add(factura);
        await db.SaveChangesAsync();
        return (await GetByIdAsync(factura.Id))!;
    }

    // ─── ESTADOS ─────────────────────────────────────────────────────────────
    public async Task<FacturaDto?> EmitirAsync(Guid id, EmitirFacturaDto dto)
    {
        var f = await db.Facturas.FindAsync(id);
        if (f is null || f.Estado != EstadoFactura.Borrador) return null;
        f.Estado = EstadoFactura.Emitida;
        if (dto.FormaPago is not null) f.FormaPago = dto.FormaPago;
        f.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    public async Task<FacturaDto?> MarcarPagadaAsync(Guid id, string? formaPago = null)
    {
        var f = await db.Facturas.FindAsync(id);
        if (f is null) return null;
        f.Estado    = EstadoFactura.Pagada;
        f.FechaPago = DateOnly.FromDateTime(DateTime.Today);
        if (formaPago is not null) f.FormaPago = formaPago;
        f.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    public async Task<FacturaDto?> AnularAsync(Guid id, AnularFacturaDto dto)
    {
        var f = await db.Facturas.FindAsync(id);
        if (f is null) return null;
        f.Estado = EstadoFactura.Anulada;
        if (!string.IsNullOrEmpty(dto.Motivo))
            f.Observaciones = $"[ANULADA: {dto.Motivo}] {f.Observaciones}";
        f.ActualizadoEn = DateTime.UtcNow;
        // Desmarcar reserva si la tiene
        if (f.ReservaId.HasValue)
        {
            var r = await db.Reservas.FindAsync(f.ReservaId.Value);
            if (r is not null) { r.Facturado = false; r.FacturaId = null; r.ActualizadoEn = DateTime.UtcNow; }
        }
        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    // ─── HTML PARA IMPRIMIR ────────────────────────────────────────────────────
    public async Task<string> GenerarHtmlAsync(Guid id)
    {
        var f = await WithIncludes(db.Facturas.Where(x => x.Id == id)).FirstOrDefaultAsync()
            ?? throw new InvalidOperationException("Factura no encontrada.");

        var lineasHtml = string.Join("", f.Lineas.OrderBy(l => l.Orden).Select(l => $@"
            <tr>
                <td style='padding:6px 8px;border-bottom:1px solid #eee;'>{l.Concepto}</td>
                <td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:center;'>{l.Cantidad} {l.Unidad}</td>
                <td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;'>{l.PrecioUnidad:F2} €</td>
                {(l.Descuento > 0 ? $"<td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;color:#c00;'>{l.Descuento:F0}%</td>" : "<td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;'>—</td>")}
                <td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;'>{l.BaseLinea:F2} €</td>
                <td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;'>{l.PorcentajeIva:F0}%</td>
                <td style='padding:6px 8px;border-bottom:1px solid #eee;text-align:right;font-weight:600;'>{l.TotalLinea:F2} €</td>
            </tr>"));

        var estadoBadge = f.Estado switch {
            EstadoFactura.Pagada  => "<span style='background:#d1fae5;color:#065f46;padding:4px 10px;border-radius:4px;font-weight:700;'>✓ PAGADA</span>",
            EstadoFactura.Anulada => "<span style='background:#fee2e2;color:#991b1b;padding:4px 10px;border-radius:4px;font-weight:700;'>✗ ANULADA</span>",
            EstadoFactura.Emitida => "<span style='background:#dbeafe;color:#1e40af;padding:4px 10px;border-radius:4px;font-weight:700;'>EMITIDA</span>",
            _ => "<span style='background:#fef3c7;color:#92400e;padding:4px 10px;border-radius:4px;font-weight:700;'>BORRADOR</span>"
        };

        return $@"<!DOCTYPE html>
<html lang='es'>
<head>
<meta charset='UTF-8'>
<title>Factura {f.NumeroFactura}</title>
<style>
  * {{ margin:0; padding:0; box-sizing:border-box; }}
  body {{ font-family: 'Segoe UI', Arial, sans-serif; font-size:13px; color:#1a1a2e; background:#fff; padding:30px; }}
  @media print {{
    body {{ padding:10px; }}
    .no-print {{ display:none !important; }}
    .page-break {{ page-break-before:always; }}
  }}
  .header {{ display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:30px; padding-bottom:20px; border-bottom:3px solid #1e3a5f; }}
  .logo {{ font-size:22px; font-weight:800; color:#1e3a5f; }}
  .logo small {{ display:block; font-size:12px; font-weight:400; color:#666; margin-top:2px; }}
  .factura-num {{ text-align:right; }}
  .factura-num h1 {{ font-size:28px; color:#1e3a5f; font-weight:900; }}
  .factura-num .fecha {{ color:#666; font-size:12px; margin-top:4px; }}
  .partes {{ display:grid; grid-template-columns:1fr 1fr; gap:30px; margin-bottom:24px; }}
  .parte h3 {{ font-size:11px; text-transform:uppercase; letter-spacing:1px; color:#888; margin-bottom:6px; }}
  .parte p {{ line-height:1.6; }}
  .parte strong {{ color:#1e3a5f; }}
  table {{ width:100%; border-collapse:collapse; margin-bottom:20px; }}
  thead tr {{ background:#1e3a5f; color:white; }}
  thead th {{ padding:8px; text-align:left; font-size:11px; text-transform:uppercase; }}
  .totales {{ display:flex; justify-content:flex-end; margin-bottom:20px; }}
  .totales-box {{ width:260px; }}
  .totales-row {{ display:flex; justify-content:space-between; padding:5px 0; border-bottom:1px solid #eee; }}
  .totales-row.total {{ font-weight:800; font-size:15px; color:#1e3a5f; border-top:2px solid #1e3a5f; border-bottom:none; padding-top:8px; }}
  .footer {{ margin-top:30px; padding-top:16px; border-top:1px solid #eee; font-size:11px; color:#999; text-align:center; }}
  .btn-print {{ display:inline-block; margin-bottom:20px; padding:10px 24px; background:#1e3a5f; color:white; border:none; border-radius:6px; font-size:13px; cursor:pointer; font-weight:600; }}
</style>
</head>
<body>
<button class='no-print btn-print' onclick='window.print()'>🖨️ Imprimir / Guardar PDF</button>
<div class='header'>
  <div class='logo'>
    {f.EmisorNombre}
    <small>{f.EmisorCif ?? ""} &nbsp;|&nbsp; {f.EmisorDireccion ?? ""} &nbsp;|&nbsp; {f.EmisorTelefono ?? ""}</small>
  </div>
  <div class='factura-num'>
    <h1>FACTURA</h1>
    <div style='font-size:18px;font-weight:700;color:#c9a84c;'>{f.NumeroFactura}</div>
    <div class='fecha'>Fecha: {f.FechaEmision:dd/MM/yyyy}</div>
    {(f.FechaVencimiento.HasValue ? $"<div class='fecha'>Vencimiento: {f.FechaVencimiento:dd/MM/yyyy}</div>" : "")}
    <div style='margin-top:6px;'>{estadoBadge}</div>
  </div>
</div>

<div class='partes'>
  <div class='parte'>
    <h3>Emisor</h3>
    <p><strong>{f.EmisorNombre}</strong><br>
    CIF: {f.EmisorCif ?? "—"}<br>
    {f.EmisorDireccion ?? ""}<br>
    {f.EmisorTelefono ?? ""}</p>
  </div>
  <div class='parte'>
    <h3>Destinatario</h3>
    <p><strong>{f.DestinatarioNombre}</strong><br>
    DNI/NIE: {f.DestinatarioDni}<br>
    {f.DestinatarioDireccion ?? ""} {f.DestinatarioCp ?? ""}<br>
    {f.DestinatarioMunicipio ?? ""}</p>
  </div>
</div>

<table>
  <thead>
    <tr>
      <th style='width:40%;'>Concepto</th>
      <th style='text-align:center;'>Cantidad</th>
      <th style='text-align:right;'>Precio/ud.</th>
      <th style='text-align:right;'>Dto.</th>
      <th style='text-align:right;'>Base</th>
      <th style='text-align:right;'>IVA</th>
      <th style='text-align:right;'>Total</th>
    </tr>
  </thead>
  <tbody>
    {lineasHtml}
  </tbody>
</table>

<div class='totales'>
  <div class='totales-box'>
    <div class='totales-row'><span>Base imponible</span><span>{f.BaseImponible:F2} €</span></div>
    <div class='totales-row'><span>IVA ({f.PorcentajeIva:F0}%)</span><span>{f.CuotaIva:F2} €</span></div>
    <div class='totales-row total'><span>TOTAL</span><span>{f.Total:F2} €</span></div>
  </div>
</div>

{(f.FormaPago is not null ? $"<p style='margin-bottom:6px;'><strong>Forma de pago:</strong> {f.FormaPago}</p>" : "")}
{(f.Observaciones is not null ? $"<p style='color:#555;font-size:12px;'><em>{f.Observaciones}</em></p>" : "")}

<div class='footer'>
  {f.EmisorNombre} &nbsp;·&nbsp; {f.EmisorCif} &nbsp;·&nbsp; Factura generada el {DateTime.Now:dd/MM/yyyy HH:mm}
</div>
</body>
</html>";
    }
}
