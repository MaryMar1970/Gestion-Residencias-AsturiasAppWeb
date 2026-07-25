using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

// ─────────────────────────────────────────────────────────────────────────────
// HUÉSPEDES
// ─────────────────────────────────────────────────────────────────────────────
public interface IHuespedService
{
    Task<List<HuespedDto>> GetAllAsync(string? buscar = null);
    Task<HuespedDto?> GetByIdAsync(Guid id);
    Task<HuespedDto?> GetByDniAsync(string dni);
    Task<HuespedDto> CreateAsync(UpsertHuespedDto dto);
    Task<HuespedDto?> UpdateAsync(Guid id, UpsertHuespedDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class HuespedService(IResidenciaDbContext db) : IHuespedService
{
    private static HuespedDto ToDto(Huesped h, int reservas = 0) => new(
        h.Id, h.Dni, h.Nombre, h.Apellidos,
        $"{h.Nombre} {h.Apellidos}",
        h.Telefono, h.Email,
        h.Direccion, h.CodigoPostal, h.Municipio, h.Provincia,
        h.CentroOrigen, h.Departamento,
        h.TipoHuesped, h.EnListaNegra, h.MotivoListaNegra,
        h.Notas, h.Empleo, h.Situacion, h.Finalidad, h.EmpleoCategoria, reservas, h.CreadoEn);

    public async Task<List<HuespedDto>> GetAllAsync(string? buscar = null)
    {
        var q = db.Huespedes.AsQueryable();
        if (!string.IsNullOrWhiteSpace(buscar))
        {
            var b = buscar.Trim().ToLower();
            q = q.Where(h =>
                h.Dni.ToLower().Contains(b) ||
                h.Nombre.ToLower().Contains(b) ||
                h.Apellidos.ToLower().Contains(b) ||
                (h.Email != null && h.Email.ToLower().Contains(b)));
        }
        return await q.OrderBy(h => h.Apellidos).ThenBy(h => h.Nombre)
            .Select(h => ToDto(h, h.Reservas.Count))
            .ToListAsync();
    }

    public async Task<HuespedDto?> GetByIdAsync(Guid id)
    {
        var h = await db.Huespedes.Include(x => x.Reservas).FirstOrDefaultAsync(x => x.Id == id);
        return h is null ? null : ToDto(h, h.Reservas.Count);
    }

    public async Task<HuespedDto?> GetByDniAsync(string dni)
    {
        var h = await db.Huespedes.Include(x => x.Reservas)
            .FirstOrDefaultAsync(x => x.Dni == dni.Trim().ToUpper());
        return h is null ? null : ToDto(h, h.Reservas.Count);
    }

    public async Task<HuespedDto> CreateAsync(UpsertHuespedDto dto)
    {
        var entity = new Huesped
        {
            Dni = dto.Dni.Trim().ToUpper(), Nombre = dto.Nombre.Trim(),
            Apellidos = dto.Apellidos.Trim(), Telefono = dto.Telefono,
            Email = dto.Email, Direccion = dto.Direccion, CodigoPostal = dto.CodigoPostal,
            Municipio = dto.Municipio, Provincia = dto.Provincia,
            CentroOrigen = dto.CentroOrigen, Departamento = dto.Departamento,
            TipoHuesped = dto.TipoHuesped, EnListaNegra = dto.EnListaNegra,
            MotivoListaNegra = dto.MotivoListaNegra, Notas = dto.Notas,
            Empleo = dto.Empleo, Situacion = dto.Situacion,
            Finalidad = dto.Finalidad, EmpleoCategoria = dto.EmpleoCategoria
        };
        db.Huespedes.Add(entity);
        await db.SaveChangesAsync();
        return ToDto(entity);
    }

    public async Task<HuespedDto?> UpdateAsync(Guid id, UpsertHuespedDto dto)
    {
        var entity = await db.Huespedes.FindAsync(id);
        if (entity is null) return null;
        entity.Dni = dto.Dni.Trim().ToUpper(); entity.Nombre = dto.Nombre.Trim();
        entity.Apellidos = dto.Apellidos.Trim(); entity.Telefono = dto.Telefono;
        entity.Email = dto.Email; entity.Direccion = dto.Direccion; entity.CodigoPostal = dto.CodigoPostal;
        entity.Municipio = dto.Municipio; entity.Provincia = dto.Provincia;
        entity.CentroOrigen = dto.CentroOrigen; entity.Departamento = dto.Departamento;
        entity.TipoHuesped = dto.TipoHuesped; entity.EnListaNegra = dto.EnListaNegra;
        entity.MotivoListaNegra = dto.MotivoListaNegra; entity.Notas = dto.Notas;
        entity.Empleo = dto.Empleo; entity.Situacion = dto.Situacion;
        entity.Finalidad = dto.Finalidad; entity.EmpleoCategoria = dto.EmpleoCategoria;
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return ToDto(entity);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Huespedes.FindAsync(id);
        if (entity is null) return false;
        db.Huespedes.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// RESERVAS
// ─────────────────────────────────────────────────────────────────────────────
public interface IReservaService
{
    Task<List<ReservaDto>> GetAllAsync(Guid? residenciaId = null, Guid? habitacionId = null,
        string? fechaDesde = null, string? fechaHasta = null, bool incluirCanceladas = false);
    Task<ReservaDto?> GetByIdAsync(Guid id);
    Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid habitacionId,
        string fechaEntrada, string fechaSalida, Guid? excluirReservaId = null, bool esBloqueo = false);
    Task<ReservaDto> CreateAsync(CrearReservaDto dto);
    Task<ReservaDto?> UpdateAsync(Guid id, ActualizarReservaDto dto);
    Task<bool> DeleteAsync(Guid id);
    Task<CalendarioDto> GetCalendarioAsync(string fechaInicio, string fechaFin, Guid? residenciaId = null);
}

public class ReservaService(IResidenciaDbContext db) : IReservaService
{
    private static ReservaDto ToDto(Reserva r) => new(
        r.Id, r.NumeroOrden,
        r.HabitacionId, r.Habitacion?.Numero ?? "", r.Habitacion?.Residencia?.Nombre ?? "", r.Habitacion?.TipoHabitacion?.Nombre ?? "",
        r.HuespedId, r.Huesped is null ? null : $"{r.Huesped.Nombre} {r.Huesped.Apellidos}", r.Huesped?.Dni, r.Huesped?.Nombre, r.Huesped?.Apellidos,
        r.Huesped?.Telefono, r.Huesped?.Email,
        r.FechaEntrada.ToString("yyyy-MM-dd"), r.FechaSalida.ToString("yyyy-MM-dd"),
        r.TotalNoches, r.NumPersonas, r.CamasSupletorias,
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

    private static IQueryable<Reserva> WithIncludes(IQueryable<Reserva> q) =>
        q.Include(r => r.Habitacion).ThenInclude(h => h.Residencia)
         .Include(r => r.Habitacion).ThenInclude(h => h.TipoHabitacion)
         .Include(r => r.Huesped);

    public async Task<List<ReservaDto>> GetAllAsync(Guid? residenciaId = null,
        Guid? habitacionId = null, string? fechaDesde = null, string? fechaHasta = null, bool incluirCanceladas = false)
    {
        var q = WithIncludes(db.Reservas.AsQueryable());
        if (!incluirCanceladas)
        {
            q = q.Where(r => r.Estado != EstadoReserva.Cancelada
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA"));
        }

        if (residenciaId.HasValue)
            q = q.Where(r => r.Habitacion.ResidenciaId == residenciaId.Value);
        if (habitacionId.HasValue)
            q = q.Where(r => r.HabitacionId == habitacionId.Value);
        if (DateOnly.TryParse(fechaDesde, out var d))
            q = q.Where(r => r.FechaSalida >= d);
        if (DateOnly.TryParse(fechaHasta, out var h))
            q = q.Where(r => r.FechaEntrada <= h);

        return await q.OrderBy(r => r.FechaEntrada).Select(r => ToDto(r)).ToListAsync();
    }

    public async Task<ReservaDto?> GetByIdAsync(Guid id)
    {
        var r = await WithIncludes(db.Reservas.Where(x => x.Id == id)).FirstOrDefaultAsync();
        return r is null ? null : ToDto(r);
    }

    public async Task<SolapamientoDto> ComprobarSolapamientoAsync(Guid habitacionId,
        string fechaEntradaStr, string fechaSalidaStr, Guid? excluirReservaId = null, bool esBloqueo = false)
    {
        var entrada = DateOnly.Parse(fechaEntradaStr);
        var salida  = DateOnly.Parse(fechaSalidaStr);

        var q = db.Reservas
            .Include(r => r.Huesped)
            .Include(r => r.Habitacion)
            .Where(r => r.HabitacionId == habitacionId
                && r.Estado != EstadoReserva.Cancelada
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA"));

        if (excluirReservaId.HasValue)
            q = q.Where(r => r.Id != excluirReservaId.Value);

        var conflictosRaw = await q.ToListAsync();

        var conflictos = conflictosRaw.Where(r =>
        {
            bool isAnyBloqueo = esBloqueo || r.EsBloqueo;
            if (isAnyBloqueo)
            {
                return entrada <= r.FechaSalida && salida >= r.FechaEntrada;
            }
            else
            {
                return entrada < r.FechaSalida && salida > r.FechaEntrada;
            }
        }).ToList();

        string? mensaje = null;
        if (conflictos.Count > 0)
        {
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

    public async Task<ReservaDto> CreateAsync(CrearReservaDto dto)
    {
        var entrada = DateOnly.Parse(dto.FechaEntrada);
        var salida  = DateOnly.Parse(dto.FechaSalida);
        if (salida <= entrada)
            throw new InvalidOperationException("La fecha de salida debe ser posterior a la de entrada.");

        // Obtener tarifa si se indicó
        Tarifa? tarifa = dto.TarifaId.HasValue
            ? await db.Tarifas.FindAsync(dto.TarifaId.Value) : null;

        Guid assignedHabitacionId = Guid.Empty;
        if (dto.HabitacionId.HasValue && dto.HabitacionId.Value != Guid.Empty)
        {
            assignedHabitacionId = dto.HabitacionId.Value;
            var solapamiento = await ComprobarSolapamientoAsync(assignedHabitacionId, dto.FechaEntrada, dto.FechaSalida, null, dto.EsBloqueo);
            if (solapamiento.HaySolapamiento)
                throw new InvalidOperationException(solapamiento.Mensaje);
        }
        else
        {
            // Asignación automática: buscar una habitación libre
            var habitacionesDisponibles = await db.Habitaciones
                .Include(h => h.TipoHabitacion)
                .Where(h => h.Activa)
                .OrderBy(h => h.Orden)
                .ToListAsync();

            var habitacionesCandidatas = habitacionesDisponibles.Where(h =>
            {
                if (tarifa != null && h.TipoHabitacionId != tarifa.TipoHabitacionId)
                {
                    return false;
                }
                var capacidadMax = h.CapacidadPersonas + (h.AdmiteSupletorias ? h.PlazasSupletorias : 0);
                if (dto.CamasSupletorias > 0 && (!h.AdmiteSupletorias || dto.CamasSupletorias > h.PlazasSupletorias))
                {
                    return false;
                }
                return capacidadMax >= dto.NumPersonas;
            }).ToList();

            foreach (var hab in habitacionesCandidatas)
            {
                var solapamiento = await ComprobarSolapamientoAsync(hab.Id, dto.FechaEntrada, dto.FechaSalida, null, dto.EsBloqueo);
                if (!solapamiento.HaySolapamiento)
                {
                    assignedHabitacionId = hab.Id;
                    break;
                }
            }

            if (assignedHabitacionId == Guid.Empty)
            {
                throw new InvalidOperationException("No se ha encontrado ninguna habitación activa y libre para las fechas y capacidad seleccionadas.");
            }
        }

        var noches = salida.DayNumber - entrada.DayNumber;
        var precioNoche = tarifa?.PrecioNoche ?? 0;
        var iva = tarifa?.PorcentajeIva ?? 10;
        var importeBase = precioNoche * noches;
        var importeIva  = importeBase * (iva / 100m);

        // Obtener huésped para calcular prioridad
        Huesped? huesped = dto.HuespedId.HasValue
            ? await db.Huespedes.FindAsync(dto.HuespedId.Value) : null;
        var (empleoCat, eval) = EvaluarReserva(huesped, dto.Finalidad);

        int nextNum = 0;
        if (!dto.EsBloqueo && dto.HuespedId.HasValue)
        {
            nextNum = await db.Reservas
                .Where(r => r.HuespedId != null && !r.EsBloqueo)
                .MaxAsync(r => (int?)r.NumeroOrden) ?? 0;
            nextNum += 1;
        }

        var entity = new Reserva
        {
            HabitacionId = assignedHabitacionId,
            HuespedId = dto.HuespedId,
            NumeroOrden = nextNum,
            FechaEntrada = entrada, FechaSalida = salida,
            NumPersonas = dto.NumPersonas, CamasSupletorias = dto.CamasSupletorias,
            EsBloqueo = dto.EsBloqueo, MotivoBloqueo = dto.MotivoBloqueo,
            TarifaId = dto.TarifaId,
            TarifaNombreSnapshot = tarifa?.NombreTarifa,
            PrecioNocheAplicado = precioNoche,
            PorcentajeIvaAplicado = iva,
            TotalNoches = noches,
            ImporteBase = importeBase, ImporteIva = importeIva,
            ImporteTotal = importeBase + importeIva,
            Estado = dto.EsBloqueo ? EstadoReserva.Bloqueada : EstadoReserva.Confirmada,
            Observaciones = dto.Observaciones,
            Finalidad = dto.Finalidad,
            Empleo = empleoCat,
            Evaluacion = eval,
            Resolucion = dto.Resolucion ?? "SI",
            FechaSolicitud = dto.FechaSolicitud
        };
        db.Reservas.Add(entity);
        await db.SaveChangesAsync();
        return (await GetByIdAsync(entity.Id))!;
    }

    public async Task<ReservaDto?> UpdateAsync(Guid id, ActualizarReservaDto dto)
    {
        var entity = await db.Reservas.FindAsync(id);
        if (entity is null) return null;

        var entrada = DateOnly.Parse(dto.FechaEntrada);
        var salida  = DateOnly.Parse(dto.FechaSalida);

        var solapamiento = await ComprobarSolapamientoAsync(entity.HabitacionId, dto.FechaEntrada, dto.FechaSalida, id, dto.EsBloqueo);
        if (solapamiento.HaySolapamiento)
            throw new InvalidOperationException(solapamiento.Mensaje);

        Tarifa? tarifa = dto.TarifaId.HasValue
            ? await db.Tarifas.FindAsync(dto.TarifaId.Value) : null;

        var noches = salida.DayNumber - entrada.DayNumber;
        var precioNoche = tarifa?.PrecioNoche ?? entity.PrecioNocheAplicado;
        var iva = tarifa?.PorcentajeIva ?? entity.PorcentajeIvaAplicado;
        var importeBase = precioNoche * noches;
        var importeIva  = importeBase * (iva / 100m);

        // Obtener huésped para calcular prioridad
        Huesped? huesped = entity.HuespedId.HasValue
            ? await db.Huespedes.FindAsync(entity.HuespedId.Value) : null;
        var (empleoCat, eval) = EvaluarReserva(huesped, dto.Finalidad);

        entity.FechaEntrada = entrada; entity.FechaSalida = salida;
        entity.NumPersonas = dto.NumPersonas; entity.CamasSupletorias = dto.CamasSupletorias;
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
        entity.Resolucion = dto.Resolucion ?? "SI";
        entity.FechaSolicitud = dto.FechaSolicitud;
        if (dto.EsBloqueo || !entity.HuespedId.HasValue)
        {
            entity.NumeroOrden = 0;
        }
        else if (entity.NumeroOrden == 0)
        {
            int nextNum = await db.Reservas
                .Where(r => r.HuespedId != null && !r.EsBloqueo)
                .MaxAsync(r => (int?)r.NumeroOrden) ?? 0;
            entity.NumeroOrden = nextNum + 1;
        }

        entity.ActualizadoEn = DateTime.UtcNow;

        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Reservas.FindAsync(id);
        if (entity is null) return false;
        // Cancelar en lugar de borrar para mantener historial
        entity.Estado = EstadoReserva.Cancelada;
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return true;
    }

    public async Task<CalendarioDto> GetCalendarioAsync(string fechaInicioStr, string fechaFinStr, Guid? residenciaId = null)
    {
        var inicio = DateOnly.Parse(fechaInicioStr);
        var fin    = DateOnly.Parse(fechaFinStr);

        var habQuery = db.Habitaciones
            .Include(h => h.Residencia)
            .Include(h => h.TipoHabitacion)
            .Where(h => h.Activa);
        if (residenciaId.HasValue)
            habQuery = habQuery.Where(h => h.ResidenciaId == residenciaId.Value);

        var habitacionesDb = await habQuery
            .OrderBy(h => h.Residencia.Orden).ThenBy(h => h.Orden)
            .ToListAsync();

        // Ordenación natural por número en memoria para evitar fallos de traducción de EF
        var habitaciones = habitacionesDb
            .OrderBy(h => h.Residencia.Orden)
            .ThenBy(h => h.Orden)
            .ThenBy(h => h.Numero, new NaturalStringComparer())
            .ToList();

        var reservas = await WithIncludes(db.Reservas
            .Where(r => r.Estado != EstadoReserva.Cancelada
                && r.FechaEntrada < fin && r.FechaSalida > inicio
                && (r.EsBloqueo || r.HuespedId == null || r.Resolucion == null || r.Resolucion == "" || r.Resolucion == "SI" || r.Resolucion == "CONCEDIDA" || r.Resolucion == "REEVALUADA")))
            .ToListAsync();

        var calendario = habitaciones.Select(h => new CalendarioHabitacionDto(
            h.Id, h.Numero, h.Residencia.Nombre, h.TipoHabitacion.Nombre, h.TipoHabitacion.Codigo,
            h.Activa, h.Orden,
            reservas.Where(r => r.HabitacionId == h.Id).Select(ToDto).ToList()
        )).ToList();

        return new CalendarioDto(fechaInicioStr, fechaFinStr, calendario);
    }

    private (string EmpleoCategoria, string Evaluacion) EvaluarReserva(Huesped? huesped, string? finalidad)
    {
        if (huesped is null) return (string.Empty, string.Empty);

        // 1. Mapear Empleo
        var e = huesped.Empleo?.Trim().ToLower() ?? string.Empty;
        string empleoCat = "GC";
        if (e.Contains("alumno"))
        {
            empleoCat = "Alumno";
        }
        else if (e.Contains("funcionario"))
        {
            empleoCat = "Funcionario en GC";
        }
        else if (e.Contains("militar en"))
        {
            empleoCat = "Militar en GC";
        }
        else if (e.Contains("militar no") || e.Contains("militar"))
        {
            empleoCat = "Militar no  GC";
        }
        else
        {
            empleoCat = "GC";
        }

        // 2. Mapear Finalidad
        var f = finalidad?.Trim().ToLower() ?? "otros";
        string finalidadMapeada = "Otros";
        if (f.Contains("comision no") || f.Contains("comisión no") || f.Contains("indem"))
        {
            finalidadMapeada = "Comisión NO indem.";
        }
        else if (f.Contains("comision") || f.Contains("comisión"))
        {
            finalidadMapeada = "Comisión";
        }
        else if (f.Contains("destino"))
        {
            finalidadMapeada = "Destino";
        }
        else if (f.Contains("enfermedad"))
        {
            finalidadMapeada = "Enfermedad";
        }
        else if (f.Contains("urgencia"))
        {
            finalidadMapeada = "Urgencia";
        }
        else if (f.Contains("sepelio"))
        {
            finalidadMapeada = "Sepelio";
        }
        else if (f.Contains("estancia"))
        {
            finalidadMapeada = "Máx. Estancia";
        }
        else
        {
            finalidadMapeada = "Otros";
        }

        // 3. Mapear Situación
        var s = huesped.Situacion?.Trim().ToLower() ?? "activo";
        string situacionMapeada = "Activo";
        if (s.Contains("viogen"))
        {
            situacionMapeada = "Viogen";
        }
        else if (s.Contains("asoc"))
        {
            situacionMapeada = "Asociación";
        }
        else if (s.Contains("reserva activo") || s.Contains("reserva activa"))
        {
            situacionMapeada = "Reserva activo";
        }
        else if (s.Contains("reserva"))
        {
            situacionMapeada = "Reserva";
        }
        else if (s.Contains("excedencia"))
        {
            situacionMapeada = "Excedencia";
        }
        else if (s.Contains("especial"))
        {
            situacionMapeada = "Especiales";
        }
        else if (s.Contains("retirado") || s.Contains("jubilado"))
        {
            situacionMapeada = "Retirado";
        }
        else if (s.Contains("viuda"))
        {
            situacionMapeada = "Viuda";
        }
        else if (s.Contains("huerfano") || s.Contains("huérfano"))
        {
            situacionMapeada = "Huerfano";
        }
        else
        {
            situacionMapeada = "Activo";
        }

        var clave = $"{finalidadMapeada.ToLower()}|{empleoCat.ToLower()}|{situacionMapeada.ToLower()}";

        var matriz = new Dictionary<string, string>
        {
            { "comisión no indem.|gc|viogen", "1, 1, 1" },
            { "comisión no indem.|gc|activo", "1, 1, 2" },
            { "comisión no indem.|gc|reserva activo", "1, 1, 2" },
            { "destino|gc|activo", "1, 1, 3" },
            { "comisión|gc|activo", "1, 1, 4" },
            { "enfermedad|gc|retirado", "2, 8, 1" },
            { "comisión|alumno|activo", "1, 2, 2" },
            { "comisión|militar en gc|activo", "1, 3, 2" },
            { "destino|militar en gc|activo", "1, 3, 2" },
            { "comisión|funcionario en gc|activo", "1, 4, 2" },
            { "destino|funcionario en gc|activo", "1, 4, 2" },
            { "enfermedad|gc|activo", "2, 1, 1" },
            { "otros|gc|viogen", "2, 1, 2" },
            { "urgencia|gc|activo", "2, 1, 3" },
            { "sepelio|gc|activo", "2, 1, 4" },
            { "máx. estancia|gc|activo", "2, 1, 5" },
            { "otros|gc|asociación", "2, 1, 6" },
            { "otros|gc|activo", "2, 1, 7" },
            { "otros|gc|reserva activo", "2, 1, 8" },
            { "otros|alumno|activo", "2, 2, 7" },
            { "otros|gc|reserva", "2, 3, 7" },
            { "otros|militar en gc|activo", "2, 4, 7" },
            { "otros|funcionario en gc|activo", "2, 5, 7" },
            { "otros|gc|excedencia", "2, 6, 7" },
            { "otros|gc|especiales", "2, 7, 7" },
            { "otros|gc|retirado", "2, 8, 7" },
            { "otros|gc|viuda", "2, 9, 7" },
            { "otros|gc|huerfano", "2, 9, 7" },
            { "otros|militar no  gc|activo", "2, 10, 7" },
            { "otros|militar no  gc|reserva", "2, 10, 7" },
            { "otros|militar no  gc|retirado", "2, 10, 7" }
        };

        if (matriz.TryGetValue(clave, out var eval))
        {
            return (empleoCat, eval);
        }

        return (empleoCat, "(NO VÁLIDO)");
    }
}

public class NaturalStringComparer : IComparer<string>
{
    public int Compare(string? x, string? y)
    {
        if (x == y) return 0;
        if (x == null) return -1;
        if (y == null) return 1;

        int ix = 0, iy = 0;
        while (ix < x.Length && iy < y.Length)
        {
            if (char.IsDigit(x[ix]) && char.IsDigit(y[iy]))
            {
                int startX = ix;
                while (ix < x.Length && char.IsDigit(x[ix])) ix++;
                int numX = int.Parse(x.Substring(startX, ix - startX));

                int startY = iy;
                while (iy < y.Length && char.IsDigit(y[iy])) iy++;
                int numY = int.Parse(y.Substring(startY, iy - startY));

                if (numX != numY) return numX.CompareTo(numY);
            }
            else
            {
                int comp = x[ix].CompareTo(y[iy]);
                if (comp != 0) return comp;
                ix++; iy++;
            }
        }
        return x.Length.CompareTo(y.Length);
    }
}
