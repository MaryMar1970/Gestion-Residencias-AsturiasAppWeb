using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public interface IDisponibilidadService
{
    Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid? habitacionId,
        string fechaEntradaStr, string fechaSalidaStr, Guid? excluirReservaId = null, bool esBloqueo = false);
    Task<List<HabitacionDto>> GetHabitacionesDisponiblesAsync(Guid residenciaId, string fechaEntrada, string fechaSalida, int pax);
    Task<CalendarioDto> GetCalendarioAsync(string fechaInicioStr, string fechaFinStr, Guid? residenciaId = null);
}

public class DisponibilidadService(IResidenciaDbContext db, ILogger<DisponibilidadService> logger) : IDisponibilidadService
{
    private static ReservaDto ToDto(Reserva r)
    {
        bool esNegativeResolucion = Resoluciones.EsNegativa(r.Resolucion);
        bool tieneAlojamientoAdjudicado = r.HabitacionId.HasValue && r.HabitacionId.Value != Guid.Empty && !esNegativeResolucion;

        return new(
            r.Id, r.NumeroOrden, r.ResidenciaId ?? r.Habitacion?.ResidenciaId,
            tieneAlojamientoAdjudicado ? r.HabitacionId : null,
            tieneAlojamientoAdjudicado ? (r.Habitacion?.Numero ?? null) : null,
            r.Residencia?.Nombre ?? (r.Habitacion?.Residencia?.Nombre ?? null),
            tieneAlojamientoAdjudicado ? (r.Habitacion?.TipoHabitacion?.Nombre ?? null) : null,
            r.HuespedId, r.Huesped is null ? null : $"{r.Huesped.Nombre} {r.Huesped.Apellidos}", r.Huesped?.Dni, r.Huesped?.Nombre, r.Huesped?.Apellidos,
            r.Huesped?.Telefono, r.Huesped?.Email,
            r.Huesped?.Situacion, r.Huesped?.Empleo, r.Huesped?.EmpleoCategoria,
            r.FechaEntrada.ToString("yyyy-MM-dd"), r.FechaSalida.ToString("yyyy-MM-dd"),
            r.TotalNoches, r.NumPersonas, r.NumNinos, r.CamasSupletorias,
            r.FamiliaNumerosa ?? "NO", r.PorcentajeDescuento,
            r.AlojamientoSolicitado,
            r.Estado.ToString(), (int)r.Estado,
            r.EsBloqueo, r.MotivoBloqueo,
            r.TarifaNombreSnapshot,
            r.PrecioNocheAplicado, r.PorcentajeIvaAplicado,
            r.ImporteBase, r.ImporteIva, r.ImporteTotal,
            r.Pagado, r.FormaPago, r.FechaPago,
            r.Facturado, r.Observaciones,
            r.Finalidad, r.Empleo, r.Evaluacion,
            r.Resolucion, r.FechaSolicitud,
            r.CreadoEn, r.ActualizadoEn);
    }

    private static IQueryable<Reserva> WithIncludes(IQueryable<Reserva> q) =>
        q.Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
         .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
         .Include(r => r.Huesped);

    public async Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid? habitacionId,
        string fechaEntradaStr, string fechaSalidaStr, Guid? excluirReservaId = null, bool esBloqueo = false)
    {
        if (!habitacionId.HasValue || habitacionId.Value == Guid.Empty)
            return new SolapamientoDto(false, null, new List<ReservaDto>());

        var entrada = DateOnly.Parse(fechaEntradaStr);
        var salida  = DateOnly.Parse(fechaSalidaStr);

        var q = db.Reservas
            .Include(r => r.Huesped)
            .Include(r => r.Habitacion)
            .Where(r => r.HabitacionId == habitacionId.Value
                && r.Estado != EstadoReserva.Cancelada
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA"));

        if (excluirReservaId.HasValue)
            q = q.Where(r => r.Id != excluirReservaId.Value);

        if (esBloqueo)
        {
            q = q.Where(r => entrada <= r.FechaSalida && salida >= r.FechaEntrada);
        }
        else
        {
            q = q.Where(r => r.EsBloqueo
                ? (entrada <= r.FechaSalida && salida >= r.FechaEntrada)
                : (entrada < r.FechaSalida && salida > r.FechaEntrada));
        }

        var conflictos = await q.ToListAsync();

        string? mensaje = null;
        if (conflictos.Count > 0)
        {
            logger.LogWarning("Solapamiento detectado para Habitación {HabitacionId} entre {Entrada} y {Salida}. Conflictos: {Cantidad}", habitacionId, fechaEntradaStr, fechaSalidaStr, conflictos.Count);
            var detalles = conflictos.Select(c =>
            {
                string tipo = c.EsBloqueo ? "BLOQUEADO" : (!c.HuespedId.HasValue ? "RESERVADO (interno)" : "RESERVADO");
                string causa = c.EsBloqueo
                    ? (c.MotivoBloqueo ?? "Mantenimiento/Reforma")
                    : (!c.HuespedId.HasValue
                        ? (c.MotivoBloqueo ?? "Reserva de habitación")
                        : $"Reserva de {c.Huesped?.Nombre} {c.Huesped?.Apellidos}");
                return $"- {tipo} del {c.FechaEntrada:dd/MM/yyyy} al {c.FechaSalida:dd/MM/yyyy}. Causa: {causa}";
            });
            mensaje = $"La habitación/apartamento se encuentra ocupada/bloqueada en las fechas seleccionadas:\n" + string.Join("\n", detalles);
        }

        return new SolapamientoDto(
            conflictos.Count > 0,
            mensaje,
            conflictos.Select(r => ToDto(r)).ToList());
    }

    public async Task<List<HabitacionDto>> GetHabitacionesDisponiblesAsync(Guid residenciaId, string fechaEntrada, string fechaSalida, int pax)
    {
        var entrada = DateOnly.Parse(fechaEntrada);
        var salida = DateOnly.Parse(fechaSalida);

        var habOcupadasIds = await db.Reservas
            .Where(r => r.HabitacionId != null
                && r.Habitacion!.ResidenciaId == residenciaId
                && r.Estado != EstadoReserva.Cancelada
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA")
                && (r.EsBloqueo ? (entrada <= r.FechaSalida && salida >= r.FechaEntrada) : (entrada < r.FechaSalida && salida > r.FechaEntrada)))
            .Select(r => r.HabitacionId!.Value)
            .Distinct()
            .ToListAsync();

        var libres = await db.Habitaciones
            .Include(h => h.Residencia)
            .Include(h => h.TipoHabitacion)
            .Where(h => h.ResidenciaId == residenciaId && h.Activa && !habOcupadasIds.Contains(h.Id))
            .OrderBy(h => h.Orden).ThenBy(h => h.Numero)
            .ToListAsync();

        return libres.Select(h => new HabitacionDto(
            h.Id, h.ResidenciaId, h.Residencia.Nombre,
            h.TipoHabitacionId, h.TipoHabitacion.Nombre, h.TipoHabitacion.Codigo,
            h.Numero, h.Nombre, h.TipoCamaPrincipal, h.CapacidadPersonas,
            h.AdmiteSupletorias, h.PlazasSupletorias,
            h.Activa, h.Notas, h.Orden)).ToList();
    }

    public async Task<CalendarioDto> GetCalendarioAsync(string fechaInicioStr, string fechaFinStr, Guid? residenciaId = null)
    {
        logger.LogInformation("GetCalendarioAsync called: fechaInicio={FI}, fechaFin={FF}, residenciaId={RId}", fechaInicioStr, fechaFinStr, residenciaId);

        var inicio = DateOnly.Parse(fechaInicioStr);
        var fin    = DateOnly.Parse(fechaFinStr);

        var habQuery = db.Habitaciones
            .Include(h => h.Residencia)
            .Include(h => h.TipoHabitacion)
            .AsQueryable();

        if (residenciaId.HasValue)
            habQuery = habQuery.Where(h => h.ResidenciaId == residenciaId.Value);

        var habitacionesDb = await habQuery.ToListAsync();

        logger.LogInformation("GetCalendarioAsync: Found {Count} habitaciones from DB (residenciaId filter={RId})", habitacionesDb.Count, residenciaId);

        var habitaciones = habitacionesDb
            .OrderBy(h => h.Residencia?.Orden ?? 0)
            .ThenBy(h => h.Orden)
            .ThenBy(h => h.Numero, new NaturalStringComparer())
            .ToList();

        var reservas = await WithIncludes(db.Reservas
            .Where(r => r.Estado != EstadoReserva.Cancelada
                && r.FechaEntrada < fin && r.FechaSalida > inicio
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA")))
            .ToListAsync();

        logger.LogInformation("GetCalendarioAsync: Found {Count} reservas in date range", reservas.Count);

        var calendario = habitaciones.Select(h => new CalendarioHabitacionDto(
            h.Id,
            h.ResidenciaId,
            h.Numero ?? "—",
            h.Residencia?.Nombre ?? "Sin Residencia",
            h.TipoHabitacion?.Nombre ?? "Sin Tipo",
            h.TipoHabitacion?.Codigo ?? "—",
            h.Activa,
            h.Orden,
            reservas.Where(r => r.HabitacionId == h.Id).Select(ToDto).ToList()
        )).ToList();

        logger.LogInformation("GetCalendarioAsync: Returning {Count} habitaciones in CalendarioDto", calendario.Count);

        return new CalendarioDto(fechaInicioStr, fechaFinStr, calendario);
    }
}
