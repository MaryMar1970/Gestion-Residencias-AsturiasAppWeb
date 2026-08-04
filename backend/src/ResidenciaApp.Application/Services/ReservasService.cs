using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public interface IReservaService
{
    Task<PaginatedResult<ReservaDto>> GetAllAsync(Guid? residenciaId = null, Guid? habitacionId = null,
        string? fechaDesde = null, string? fechaHasta = null, bool incluirCanceladas = false,
        int page = 1, int pageSize = 50);
    Task<ReservaDto?> GetByIdAsync(Guid id);
    Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid? habitacionId,
        string fechaEntrada, string fechaSalida, Guid? excluirReservaId = null, bool esBloqueo = false);
    Task<List<HabitacionDto>> GetHabitacionesDisponiblesAsync(Guid residenciaId, string fechaEntrada, string fechaSalida, int pax);
    Task<ResultadoOperacionReservaDto> CreateAsync(CrearReservaDto dto);
    Task<ResultadoOperacionReservaDto?> UpdateAsync(Guid id, ActualizarReservaDto dto);
    Task<ResultadoOperacionReservaDto?> MoverAsync(Guid id, MoverReservaDto dto);
    Task<List<ReservaDto>> GetCandidatosReevaluacionAsync(Guid reservaId);
    Task<bool> DeleteAsync(Guid id);
    Task<CalendarioDto> GetCalendarioAsync(string fechaInicio, string fechaFin, Guid? residenciaId = null);
}

public class ReservaService(IResidenciaDbContext db, IEvaluacionService evalSvc, IDisponibilidadService dispSvc, ILogger<ReservaService> logger) : IReservaService
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
        q.Include(r => r.Residencia)
         .Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
         .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
         .Include(r => r.Huesped);

    public async Task<PaginatedResult<ReservaDto>> GetAllAsync(Guid? residenciaId = null,
        Guid? habitacionId = null, string? fechaDesde = null, string? fechaHasta = null, bool incluirCanceladas = false,
        int page = 1, int pageSize = 50)
    {
        page = Math.Max(1, page);
        pageSize = Math.Clamp(pageSize, 1, 500);

        var q = WithIncludes(db.Reservas.AsQueryable());
        if (!incluirCanceladas)
        {
            q = q.Where(r => r.Estado != EstadoReserva.Cancelada
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA"));
        }

        if (residenciaId.HasValue)
            q = q.Where(r => r.ResidenciaId == residenciaId.Value || (r.Habitacion != null && r.Habitacion.ResidenciaId == residenciaId.Value));
        if (habitacionId.HasValue)
            q = q.Where(r => r.HabitacionId == habitacionId.Value);
        if (DateOnly.TryParse(fechaDesde, out var d))
            q = q.Where(r => r.FechaSalida >= d);
        if (DateOnly.TryParse(fechaHasta, out var h))
            q = q.Where(r => r.FechaEntrada <= h);

        var totalCount = await q.CountAsync();
        var items = await q.OrderBy(r => r.FechaEntrada)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .Select(r => ToDto(r))
            .ToListAsync();

        int totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
        return new PaginatedResult<ReservaDto>(items, totalCount, page, pageSize, totalPages);
    }

    public async Task<ReservaDto?> GetByIdAsync(Guid id)
    {
        var r = await WithIncludes(db.Reservas.Where(x => x.Id == id)).FirstOrDefaultAsync();
        return r is null ? null : ToDto(r);
    }

    public Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid? habitacionId,
        string fechaEntradaStr, string fechaSalidaStr, Guid? excluirReservaId = null, bool esBloqueo = false)
    {
        return dispSvc.ComprobarSolapamientoAsync(habitacionId, fechaEntradaStr, fechaSalidaStr, excluirReservaId, esBloqueo);
    }

    public Task<List<HabitacionDto>> GetHabitacionesDisponiblesAsync(Guid residenciaId, string fechaEntrada, string fechaSalida, int pax)
    {
        return dispSvc.GetHabitacionesDisponiblesAsync(residenciaId, fechaEntrada, fechaSalida, pax);
    }

    public Task<CalendarioDto> GetCalendarioAsync(string fechaInicioStr, string fechaFinStr, Guid? residenciaId = null)
    {
        return dispSvc.GetCalendarioAsync(fechaInicioStr, fechaFinStr, residenciaId);
    }

    public async Task<List<ReservaDto>> GetCandidatosReevaluacionAsync(Guid reservaId)
    {
        var reservaRenuncia = await db.Reservas
            .Include(r => r.Habitacion)
            .FirstOrDefaultAsync(r => r.Id == reservaId);

        if (reservaRenuncia is null) return new List<ReservaDto>();

        var habitacion = reservaRenuncia.Habitacion;
        if (habitacion is null) return new List<ReservaDto>();

        var residenciaId = habitacion.ResidenciaId;
        var inicio = reservaRenuncia.FechaEntrada;
        var fin = reservaRenuncia.FechaSalida;

        var candidatas = await db.Reservas
            .Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
            .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
            .Include(r => r.Huesped)
            .Where(r => r.Habitacion.ResidenciaId == residenciaId
                && r.Id != reservaRenuncia.Id
                && r.Estado != EstadoReserva.Cancelada
                && (r.Resolucion == Resoluciones.No || r.Resolucion == Resoluciones.Denegada)
                && r.FechaEntrada < fin && r.FechaSalida > inicio)
            .OrderBy(r => r.FechaSolicitud ?? r.CreadoEn)
            .ThenBy(r => r.NumeroOrden)
            .ToListAsync();

        var result = new List<ReservaDto>();
        foreach (var cand in candidatas)
        {
            var solapamiento = await dispSvc.ComprobarSolapamientoAsync(
                cand.HabitacionId,
                cand.FechaEntrada.ToString("yyyy-MM-dd"),
                cand.FechaSalida.ToString("yyyy-MM-dd"),
                cand.Id);

            if (!solapamiento.HaySolapamiento)
            {
                result.Add(ToDto(cand));
            }
        }

        return result;
    }

    private async Task<List<ReservaDto>> ReevaluarCandidatosTrasRenunciaAsync(Reserva reservaRenuncia)
    {
        var reevaluadas = new List<ReservaDto>();
        var habitacion = await db.Habitaciones
            .Include(h => h.Residencia)
            .FirstOrDefaultAsync(h => h.Id == reservaRenuncia.HabitacionId);

        if (habitacion is null) return reevaluadas;

        var residenciaId = habitacion.ResidenciaId;
        var inicio = reservaRenuncia.FechaEntrada;
        var fin = reservaRenuncia.FechaSalida;

        var candidatas = await db.Reservas
            .Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
            .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
            .Include(r => r.Huesped)
            .Where(r => r.Habitacion.ResidenciaId == residenciaId
                && r.Id != reservaRenuncia.Id
                && r.Estado != EstadoReserva.Cancelada
                && (r.Resolucion == Resoluciones.No || r.Resolucion == Resoluciones.Denegada)
                && r.FechaEntrada < fin && r.FechaSalida > inicio)
            .OrderBy(r => r.FechaSolicitud ?? r.CreadoEn)
            .ThenBy(r => r.NumeroOrden)
            .ToListAsync();

        foreach (var cand in candidatas)
        {
            var solapamiento = await dispSvc.ComprobarSolapamientoAsync(
                cand.HabitacionId,
                cand.FechaEntrada.ToString("yyyy-MM-dd"),
                cand.FechaSalida.ToString("yyyy-MM-dd"),
                cand.Id);

            if (!solapamiento.HaySolapamiento)
            {
                cand.Resolucion = Resoluciones.Reevaluada;
                cand.ActualizadoEn = DateTime.UtcNow;
                var candDto = await GetByIdAsync(cand.Id);
                if (candDto is not null) reevaluadas.Add(candDto);
            }
        }

        if (reevaluadas.Count > 0)
        {
            await db.SaveChangesAsync();
        }

        return reevaluadas;
    }

    public async Task<ResultadoOperacionReservaDto> CreateAsync(CrearReservaDto dto)
    {
        var entrada = DateOnly.Parse(dto.FechaEntrada);
        var salida  = DateOnly.Parse(dto.FechaSalida);
        if (salida <= entrada)
            throw new InvalidOperationException("La fecha de salida debe ser posterior a la de entrada.");

        Tarifa? tarifa = dto.TarifaId.HasValue
            ? await db.Tarifas.FindAsync(dto.TarifaId.Value) : null;

        Guid? assignedHabitacionId = null;
        if (dto.HabitacionId.HasValue && dto.HabitacionId.Value != Guid.Empty)
        {
            assignedHabitacionId = dto.HabitacionId.Value;
            var solapamiento = await dispSvc.ComprobarSolapamientoAsync(assignedHabitacionId, dto.FechaEntrada, dto.FechaSalida, null, dto.EsBloqueo);
            if (solapamiento.HaySolapamiento)
                throw new InvalidOperationException(solapamiento.Mensaje);
        }
        else
        {
            var habitacionesDisponibles = await db.Habitaciones
                .Include(h => h.TipoHabitacion)
                .Where(h => h.Activa)
                .OrderBy(h => h.Orden)
                .ToListAsync();

            var habitacionesCandidatas = habitacionesDisponibles.Where(h =>
            {
                if (tarifa != null && h.TipoHabitacionId != tarifa.TipoHabitacionId) return false;
                var capacidadMax = h.CapacidadPersonas + (h.AdmiteSupletorias ? h.PlazasSupletorias : 0);
                if (dto.CamasSupletorias > 0 && (!h.AdmiteSupletorias || dto.CamasSupletorias > h.PlazasSupletorias)) return false;
                return capacidadMax >= dto.NumPersonas;
            }).ToList();

            foreach (var hab in habitacionesCandidatas)
            {
                var solapamiento = await dispSvc.ComprobarSolapamientoAsync(hab.Id, dto.FechaEntrada, dto.FechaSalida, null, dto.EsBloqueo);
                if (!solapamiento.HaySolapamiento)
                {
                    assignedHabitacionId = hab.Id;
                    break;
                }
            }

            if (!assignedHabitacionId.HasValue && dto.HabitacionId.HasValue && dto.HabitacionId.Value != Guid.Empty && !Resoluciones.EsNegativa(dto.Resolucion))
            {
                throw new InvalidOperationException("No se ha encontrado ninguna habitación activa y libre para las fechas y capacidad seleccionadas.");
            }
        }

        var noches = salida.DayNumber - entrada.DayNumber;
        var precioNoche = tarifa?.PrecioNoche ?? 0;
        var iva = tarifa?.PorcentajeIva ?? 10;
        var importeBase = precioNoche * noches;
        var importeIva  = importeBase * (iva / 100m);

        Huesped? huesped = dto.HuespedId.HasValue
            ? await db.Huespedes.FindAsync(dto.HuespedId.Value) : null;
        var (empleoCat, eval) = evalSvc.EvaluarReserva(huesped, dto.Finalidad);

        Guid? targetResidenciaId = dto.ResidenciaId;
        if (!targetResidenciaId.HasValue && assignedHabitacionId.HasValue)
        {
            var habObj = await db.Habitaciones.FindAsync(assignedHabitacionId.Value);
            targetResidenciaId = habObj?.ResidenciaId;
        }

        int nextNum = 0;
        if (!dto.EsBloqueo && dto.HuespedId.HasValue)
        {
            if (dto.NumeroOrden.HasValue && dto.NumeroOrden.Value > 0)
            {
                nextNum = dto.NumeroOrden.Value;
            }
            else
            {
                nextNum = await db.Reservas
                    .Where(r => (r.ResidenciaId == targetResidenciaId || (r.Habitacion != null && r.Habitacion.ResidenciaId == targetResidenciaId))
                        && r.HuespedId != null && !r.EsBloqueo && r.Estado != EstadoReserva.Cancelada)
                    .MaxAsync(r => (int?)r.NumeroOrden) ?? 0;
                nextNum += 1;
            }
        }

        var fnStr = dto.FamiliaNumerosa ?? huesped?.FamiliaNumerosa ?? "NO";
        var descPct = dto.PorcentajeDescuento.HasValue
            ? dto.PorcentajeDescuento.Value
            : (huesped?.PorcentajeDescuento ?? (fnStr == "ESPECIAL" ? 50m : (fnStr == "GENERAL" ? 20m : 0m)));

        string resolucionCalculada;
        if (!string.IsNullOrWhiteSpace(dto.Resolucion))
        {
            resolucionCalculada = dto.Resolucion;
        }
        else
        {
            bool tieneAlojamiento = assignedHabitacionId.HasValue && assignedHabitacionId.Value != Guid.Empty;
            resolucionCalculada = DisponibilidadService_CalcularResolucionAutomatica(dto.FechaSolicitud, entrada, salida, tieneAlojamiento);
        }

        if (Resoluciones.EsNegativa(resolucionCalculada))
        {
            assignedHabitacionId = null;
        }

        var entity = new Reserva
        {
            ResidenciaId = targetResidenciaId,
            HabitacionId = assignedHabitacionId,
            HuespedId = dto.HuespedId,
            NumeroOrden = dto.EsBloqueo || !dto.HuespedId.HasValue ? 0 : nextNum,
            FechaEntrada = entrada, FechaSalida = salida,
            NumPersonas = dto.NumPersonas, NumNinos = dto.NumNinos, CamasSupletorias = dto.CamasSupletorias,
            FamiliaNumerosa = fnStr, PorcentajeDescuento = descPct,
            AlojamientoSolicitado = dto.AlojamientoSolicitado,
            EsBloqueo = dto.EsBloqueo, MotivoBloqueo = dto.MotivoBloqueo,
            TarifaId = dto.TarifaId,
            TarifaNombreSnapshot = tarifa?.NombreTarifa,
            PrecioNocheAplicado = precioNoche, PorcentajeIvaAplicado = iva,
            TotalNoches = noches,
            ImporteBase = importeBase, ImporteIva = importeIva, ImporteTotal = importeBase + importeIva,
            Estado = dto.EsBloqueo ? EstadoReserva.Bloqueada : EstadoReserva.Confirmada,
            Pagado = false, FormaPago = null, FechaPago = null,
            Observaciones = dto.Observaciones,
            Finalidad = dto.Finalidad, Empleo = empleoCat, Evaluacion = eval,
            Resolucion = resolucionCalculada, FechaSolicitud = dto.FechaSolicitud
        };

        using var transaction = await db.Database.BeginTransactionAsync();
        db.Reservas.Add(entity);
        await db.SaveChangesAsync();
        await transaction.CommitAsync();

        logger.LogInformation("Reserva {ReservaId} (Orden {NumeroOrden}) creada correctamente para Huésped {HuespedId} en Habitación {HabitacionId}", entity.Id, entity.NumeroOrden, entity.HuespedId, entity.HabitacionId);
        var creada = (await GetByIdAsync(entity.Id))!;
        return new ResultadoOperacionReservaDto(creada, new List<ReservaDto>());
    }

    public async Task<ResultadoOperacionReservaDto?> UpdateAsync(Guid id, ActualizarReservaDto dto)
    {
        var entity = await db.Reservas.FindAsync(id);
        if (entity is null)
        {
            logger.LogWarning("Intento de actualizar reserva inexistente {ReservaId}", id);
            return null;
        }

        var entrada = DateOnly.Parse(dto.FechaEntrada);
        var salida  = DateOnly.Parse(dto.FechaSalida);

        string resolucionAnterior = entity.Resolucion;

        Guid? targetHabitacionId = dto.HabitacionId.HasValue
            ? (dto.HabitacionId.Value == Guid.Empty ? null : dto.HabitacionId.Value)
            : entity.HabitacionId;

        Tarifa? tarifa = dto.TarifaId.HasValue
            ? await db.Tarifas.FindAsync(dto.TarifaId.Value) : null;

        var noches = salida.DayNumber - entrada.DayNumber;
        var precioNoche = tarifa?.PrecioNoche ?? entity.PrecioNocheAplicado;
        var iva = tarifa?.PorcentajeIva ?? entity.PorcentajeIvaAplicado;
        var importeBase = precioNoche * noches;
        var importeIva  = importeBase * (iva / 100m);

        Huesped? huesped = entity.HuespedId.HasValue
            ? await db.Huespedes.FindAsync(entity.HuespedId.Value) : null;

        if (huesped is not null)
        {
            if (!string.IsNullOrWhiteSpace(dto.Empleo)) huesped.Empleo = dto.Empleo;
            if (!string.IsNullOrWhiteSpace(dto.Situacion)) huesped.Situacion = dto.Situacion;
        }

        var (empleoCat, eval) = evalSvc.EvaluarReserva(huesped, dto.Finalidad, dto.Empleo, dto.Situacion);

        entity.FechaEntrada = entrada; entity.FechaSalida = salida;
        entity.NumPersonas = dto.NumPersonas; entity.NumNinos = dto.NumNinos; entity.CamasSupletorias = dto.CamasSupletorias;
        if (dto.FamiliaNumerosa != null) entity.FamiliaNumerosa = dto.FamiliaNumerosa;
        if (dto.PorcentajeDescuento.HasValue) entity.PorcentajeDescuento = dto.PorcentajeDescuento.Value;
        if (dto.AlojamientoSolicitado != null) entity.AlojamientoSolicitado = dto.AlojamientoSolicitado;
        entity.EsBloqueo = dto.EsBloqueo; entity.MotivoBloqueo = dto.MotivoBloqueo;
        entity.TarifaId = dto.TarifaId;
        if (tarifa is not null) entity.TarifaNombreSnapshot = tarifa.NombreTarifa;
        entity.PrecioNocheAplicado = precioNoche; entity.PorcentajeIvaAplicado = iva;
        entity.TotalNoches = noches;
        entity.ImporteBase = importeBase; entity.ImporteIva = importeIva;
        entity.ImporteTotal = importeBase + importeIva;
        entity.Estado = Enum.Parse<EstadoReserva>(dto.Estado);
        entity.Pagado = dto.Pagado; entity.FormaPago = dto.FormaPago;
        if (dto.Pagado && entity.FechaPago is null) entity.FechaPago = DateTime.UtcNow;
        entity.Observaciones = dto.Observaciones;
        entity.Finalidad = dto.Finalidad;
        entity.Empleo = empleoCat;
        entity.Evaluacion = eval;

        string resolucionFinal;
        if (!string.IsNullOrWhiteSpace(dto.Resolucion))
        {
            resolucionFinal = dto.Resolucion;
        }
        else
        {
            bool tieneAlojamiento = targetHabitacionId.HasValue && targetHabitacionId.Value != Guid.Empty;
            resolucionFinal = DisponibilidadService_CalcularResolucionAutomatica(dto.FechaSolicitud ?? entity.FechaSolicitud, entrada, salida, tieneAlojamiento);
        }

        bool esNegativeResolucion = Resoluciones.EsNegativa(resolucionFinal);
        if (esNegativeResolucion)
        {
            targetHabitacionId = null;
        }

        if (targetHabitacionId.HasValue && targetHabitacionId.Value != Guid.Empty)
        {
            var solapamiento = await dispSvc.ComprobarSolapamientoAsync(targetHabitacionId.Value, dto.FechaEntrada, dto.FechaSalida, id, dto.EsBloqueo);
            if (solapamiento.HaySolapamiento)
                throw new InvalidOperationException(solapamiento.Mensaje);
        }

        entity.HabitacionId = targetHabitacionId;
        entity.Resolucion = resolucionFinal;
        entity.FechaSolicitud = dto.FechaSolicitud;
        if (entity.Estado == EstadoReserva.Cancelada || dto.EsBloqueo || !entity.HuespedId.HasValue)
        {
            entity.NumeroOrden = 0;
        }
        else if (entity.NumeroOrden == 0)
        {
            int nextNum = await db.Reservas
                .Where(r => r.HuespedId != null && !r.EsBloqueo && r.Estado != EstadoReserva.Cancelada)
                .MaxAsync(r => (int?)r.NumeroOrden) ?? 0;
            entity.NumeroOrden = nextNum + 1;
        }

        using var transaction = await db.Database.BeginTransactionAsync();
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();

        var reevaluadas = new List<ReservaDto>();
        bool eraPositiva = Resoluciones.EsPositivaOValida(resolucionAnterior);
        if (eraPositiva && resolucionFinal == Resoluciones.Renuncia)
        {
            if (dto.ReevaluarCandidatoId.HasValue && dto.ReevaluarCandidatoId.Value != Guid.Empty)
            {
                var candTarget = await db.Reservas
                    .Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
                    .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
                    .Include(r => r.Huesped)
                    .FirstOrDefaultAsync(r => r.Id == dto.ReevaluarCandidatoId.Value);

                if (candTarget is not null && (candTarget.Resolucion == Resoluciones.No || candTarget.Resolucion == Resoluciones.Denegada))
                {
                    candTarget.Resolucion = Resoluciones.Reevaluada;
                    candTarget.ActualizadoEn = DateTime.UtcNow;
                    await db.SaveChangesAsync();
                    var candDto = (await GetByIdAsync(candTarget.Id))!;
                    reevaluadas.Add(candDto);
                }
            }
        }

        await transaction.CommitAsync();
        logger.LogInformation("Reserva {ReservaId} actualizada correctamente con Resolución {Resolucion}", id, entity.Resolucion);

        var actualizada = (await GetByIdAsync(id))!;
        return new ResultadoOperacionReservaDto(actualizada, reevaluadas);
    }

    public async Task<ResultadoOperacionReservaDto?> MoverAsync(Guid id, MoverReservaDto dto)
    {
        var entity = await db.Reservas.FindAsync(id);
        if (entity is null)
        {
            logger.LogWarning("Intento de mover reserva inexistente {ReservaId}", id);
            return null;
        }

        var entrada = DateOnly.Parse(dto.FechaEntrada);
        var salida  = DateOnly.Parse(dto.FechaSalida);

        if (salida <= entrada)
            throw new InvalidOperationException("La fecha de salida debe ser posterior a la de entrada.");

        var solapamiento = await dispSvc.ComprobarSolapamientoAsync(dto.HabitacionId, dto.FechaEntrada, dto.FechaSalida, id, entity.EsBloqueo);
        if (solapamiento.HaySolapamiento)
            throw new InvalidOperationException(solapamiento.Mensaje);

        var noches = salida.DayNumber - entrada.DayNumber;
        var importeBase = entity.PrecioNocheAplicado * noches;
        var importeIva  = importeBase * (entity.PorcentajeIvaAplicado / 100m);

        entity.HabitacionId = dto.HabitacionId;
        entity.FechaEntrada = entrada;
        entity.FechaSalida  = salida;
        entity.TotalNoches  = noches;
        entity.ImporteBase  = importeBase;
        entity.ImporteIva   = importeIva;
        entity.ImporteTotal = importeBase + importeIva;
        entity.ActualizadoEn = DateTime.UtcNow;

        using var transaction = await db.Database.BeginTransactionAsync();
        await db.SaveChangesAsync();
        await transaction.CommitAsync();
        var actualizada = (await GetByIdAsync(entity.Id))!;
        return new ResultadoOperacionReservaDto(actualizada, new List<ReservaDto>());
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Reservas.FindAsync(id);
        if (entity is null)
        {
            logger.LogWarning("Intento de eliminar reserva inexistente {ReservaId}", id);
            return false;
        }

        logger.LogInformation("Eliminando reserva {ReservaId}", id);
        db.Reservas.Remove(entity);
        await db.SaveChangesAsync();
        logger.LogInformation("Reserva {ReservaId} eliminada correctamente", id);
        return true;
    }

    private static string DisponibilidadService_CalcularResolucionAutomatica(DateTime? fechaSolicitud, DateOnly fechaEntrada, DateOnly fechaSalida, bool tieneAlojamiento)
    {
        var fechaSol = fechaSolicitud?.Date ?? DateTime.Today;
        var entradaDt = fechaEntrada.ToDateTime(TimeOnly.MinValue);
        var salidaDt  = fechaSalida.ToDateTime(TimeOnly.MinValue);

        int totalNoches = (salidaDt.Date - entradaDt.Date).Days;
        int diasAntelacion = (entradaDt.Date - fechaSol.Date).Days;

        if (totalNoches > 7 || diasAntelacion > 30) return Resoluciones.Desestimada;

        int dayOfWeek = (int)fechaSol.DayOfWeek;
        int dayOfWeekMondayBased = dayOfWeek == 0 ? 7 : dayOfWeek;
        int umbral = 14 - dayOfWeekMondayBased;

        bool esProximidad = diasAntelacion <= umbral;
        if (tieneAlojamiento) return esProximidad ? Resoluciones.Concedida : Resoluciones.Si;
        else return esProximidad ? Resoluciones.Denegada : Resoluciones.No;
    }
}
