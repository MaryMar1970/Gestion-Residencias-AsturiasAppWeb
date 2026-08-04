using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

// ─────────────────────────────────────────────────────────────────────────────
// RESIDENCIAS
// ─────────────────────────────────────────────────────────────────────────────
public interface IResidenciaService
{
    Task<List<ResidenciaDto>> GetAllAsync();
    Task<ResidenciaDto?> GetByIdAsync(Guid id);
    Task<ResidenciaDto> CreateAsync(UpsertResidenciaDto dto);
    Task<ResidenciaDto?> UpdateAsync(Guid id, UpsertResidenciaDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class ResidenciaService(IResidenciaDbContext db) : IResidenciaService
{
    public async Task<List<ResidenciaDto>> GetAllAsync() =>
        await db.Residencias
            .OrderBy(r => r.Orden).ThenBy(r => r.Nombre)
            .Select(r => new ResidenciaDto(
                r.Id, r.Nombre, r.RazonSocial, r.Cif,
                r.Direccion, r.CodigoPostal, r.Municipio, r.Provincia,
                r.Telefono, r.Email, r.SerieFactura, r.Activa, r.Orden,
                r.Habitaciones.Count(h => h.Activa)))
            .ToListAsync();

    public async Task<ResidenciaDto?> GetByIdAsync(Guid id) =>
        await db.Residencias
            .Where(r => r.Id == id)
            .Select(r => new ResidenciaDto(
                r.Id, r.Nombre, r.RazonSocial, r.Cif,
                r.Direccion, r.CodigoPostal, r.Municipio, r.Provincia,
                r.Telefono, r.Email, r.SerieFactura, r.Activa, r.Orden,
                r.Habitaciones.Count(h => h.Activa)))
            .FirstOrDefaultAsync();

    public async Task<ResidenciaDto> CreateAsync(UpsertResidenciaDto dto)
    {
        var entity = new Residencia
        {
            Nombre = dto.Nombre, RazonSocial = dto.RazonSocial, Cif = dto.Cif,
            Direccion = dto.Direccion, CodigoPostal = dto.CodigoPostal,
            Municipio = dto.Municipio, Provincia = dto.Provincia,
            Telefono = dto.Telefono, Email = dto.Email, SerieFactura = dto.SerieFactura,
            Activa = dto.Activa, Orden = dto.Orden
        };
        db.Residencias.Add(entity);
        await db.SaveChangesAsync();
        return new ResidenciaDto(entity.Id, entity.Nombre, entity.RazonSocial, entity.Cif,
            entity.Direccion, entity.CodigoPostal, entity.Municipio, entity.Provincia,
            entity.Telefono, entity.Email, entity.SerieFactura, entity.Activa, entity.Orden, 0);
    }

    public async Task<ResidenciaDto?> UpdateAsync(Guid id, UpsertResidenciaDto dto)
    {
        var entity = await db.Residencias.FindAsync(id);
        if (entity is null) return null;
        entity.Nombre = dto.Nombre; entity.RazonSocial = dto.RazonSocial; entity.Cif = dto.Cif;
        entity.Direccion = dto.Direccion; entity.CodigoPostal = dto.CodigoPostal;
        entity.Municipio = dto.Municipio; entity.Provincia = dto.Provincia;
        entity.Telefono = dto.Telefono; entity.Email = dto.Email; entity.SerieFactura = dto.SerieFactura;
        entity.Activa = dto.Activa; entity.Orden = dto.Orden;
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        var total = await db.Habitaciones.CountAsync(h => h.ResidenciaId == id && h.Activa);
        return new ResidenciaDto(entity.Id, entity.Nombre, entity.RazonSocial, entity.Cif,
            entity.Direccion, entity.CodigoPostal, entity.Municipio, entity.Provincia,
            entity.Telefono, entity.Email, entity.SerieFactura, entity.Activa, entity.Orden, total);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Residencias.FindAsync(id);
        if (entity is null) return false;
        db.Residencias.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// TIPOS DE HABITACIÓN
// ─────────────────────────────────────────────────────────────────────────────
public interface ITipoHabitacionService
{
    Task<List<TipoHabitacionDto>> GetAllAsync();
    Task<TipoHabitacionDto?> GetByIdAsync(Guid id);
    Task<TipoHabitacionDto> CreateAsync(UpsertTipoHabitacionDto dto);
    Task<TipoHabitacionDto?> UpdateAsync(Guid id, UpsertTipoHabitacionDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class TipoHabitacionService(IResidenciaDbContext db) : ITipoHabitacionService
{
    public async Task<List<TipoHabitacionDto>> GetAllAsync() =>
        await db.TiposHabitacion.Where(t => t.Activo)
            .OrderBy(t => t.Orden).ThenBy(t => t.Nombre)
            .Select(t => new TipoHabitacionDto(t.Id, t.Nombre, t.Codigo,
                t.CapacidadMaxima, t.AdmiteSupletorias, t.Descripcion, t.Activo, t.Orden))
            .ToListAsync();

    public async Task<TipoHabitacionDto?> GetByIdAsync(Guid id) =>
        await db.TiposHabitacion.Where(t => t.Id == id)
            .Select(t => new TipoHabitacionDto(t.Id, t.Nombre, t.Codigo,
                t.CapacidadMaxima, t.AdmiteSupletorias, t.Descripcion, t.Activo, t.Orden))
            .FirstOrDefaultAsync();

    public async Task<TipoHabitacionDto> CreateAsync(UpsertTipoHabitacionDto dto)
    {
        var entity = new TipoHabitacion
        {
            Nombre = dto.Nombre, Codigo = dto.Codigo.ToUpper(),
            CapacidadMaxima = dto.CapacidadMaxima, AdmiteSupletorias = dto.AdmiteSupletorias,
            Descripcion = dto.Descripcion, Activo = dto.Activo, Orden = dto.Orden
        };
        db.TiposHabitacion.Add(entity);
        await db.SaveChangesAsync();
        return new TipoHabitacionDto(entity.Id, entity.Nombre, entity.Codigo,
            entity.CapacidadMaxima, entity.AdmiteSupletorias, entity.Descripcion, entity.Activo, entity.Orden);
    }

    public async Task<TipoHabitacionDto?> UpdateAsync(Guid id, UpsertTipoHabitacionDto dto)
    {
        var entity = await db.TiposHabitacion.FindAsync(id);
        if (entity is null) return null;
        entity.Nombre = dto.Nombre; entity.Codigo = dto.Codigo.ToUpper();
        entity.CapacidadMaxima = dto.CapacidadMaxima; entity.AdmiteSupletorias = dto.AdmiteSupletorias;
        entity.Descripcion = dto.Descripcion; entity.Activo = dto.Activo; entity.Orden = dto.Orden;
        await db.SaveChangesAsync();
        return new TipoHabitacionDto(entity.Id, entity.Nombre, entity.Codigo,
            entity.CapacidadMaxima, entity.AdmiteSupletorias, entity.Descripcion, entity.Activo, entity.Orden);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.TiposHabitacion.FindAsync(id);
        if (entity is null) return false;
        db.TiposHabitacion.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// HABITACIONES
// ─────────────────────────────────────────────────────────────────────────────
public interface IHabitacionService
{
    Task<List<HabitacionDto>> GetAllAsync(Guid? residenciaId = null);
    Task<HabitacionDto?> GetByIdAsync(Guid id);
    Task<HabitacionDto> CreateAsync(UpsertHabitacionDto dto);
    Task<List<HabitacionDto>> CreateBatchAsync(CrearLoteHabitacionesDto dto);
    Task<HabitacionDto?> UpdateAsync(Guid id, UpsertHabitacionDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class HabitacionService(IResidenciaDbContext db) : IHabitacionService
{
    public async Task<List<HabitacionDto>> GetAllAsync(Guid? residenciaId = null)
    {
        var query = db.Habitaciones
            .Include(h => h.Residencia)
            .Include(h => h.TipoHabitacion)
            .AsQueryable();
        if (residenciaId.HasValue)
            query = query.Where(h => h.ResidenciaId == residenciaId.Value);
        return await query
            .OrderBy(h => h.Residencia.Orden).ThenBy(h => h.Orden).ThenBy(h => h.Numero)
            .Select(h => new HabitacionDto(
                h.Id, h.ResidenciaId, h.Residencia.Nombre,
                h.TipoHabitacionId, h.TipoHabitacion.Nombre, h.TipoHabitacion.Codigo,
                h.Numero, h.Nombre, h.TipoCamaPrincipal, h.CapacidadPersonas,
                h.AdmiteSupletorias, h.PlazasSupletorias,
                h.Activa, h.Notas, h.Orden))
            .ToListAsync();
    }

    public async Task<HabitacionDto?> GetByIdAsync(Guid id) =>
        await db.Habitaciones
            .Include(h => h.Residencia).Include(h => h.TipoHabitacion)
            .Where(h => h.Id == id)
            .Select(h => new HabitacionDto(
                h.Id, h.ResidenciaId, h.Residencia.Nombre,
                h.TipoHabitacionId, h.TipoHabitacion.Nombre, h.TipoHabitacion.Codigo,
                h.Numero, h.Nombre, h.TipoCamaPrincipal, h.CapacidadPersonas,
                h.AdmiteSupletorias, h.PlazasSupletorias,
                h.Activa, h.Notas, h.Orden))
            .FirstOrDefaultAsync();

    public async Task<HabitacionDto> CreateAsync(UpsertHabitacionDto dto)
    {
        var entity = new Habitacion
        {
            ResidenciaId = dto.ResidenciaId, TipoHabitacionId = dto.TipoHabitacionId,
            Numero = dto.Numero, Nombre = dto.Nombre,
            TipoCamaPrincipal = string.IsNullOrWhiteSpace(dto.TipoCamaPrincipal) ? "IND" : dto.TipoCamaPrincipal,
            CapacidadPersonas = dto.CapacidadPersonas,
            AdmiteSupletorias = dto.AdmiteSupletorias, PlazasSupletorias = dto.PlazasSupletorias,
            Activa = dto.Activa, Notas = dto.Notas, Orden = dto.Orden
        };
        db.Habitaciones.Add(entity);
        await db.SaveChangesAsync();
        return (await GetByIdAsync(entity.Id))!;
    }

    public async Task<List<HabitacionDto>> CreateBatchAsync(CrearLoteHabitacionesDto dto)
    {
        var creadas = new List<Habitacion>();
        for (int i = 0; i < dto.Cantidad; i++)
        {
            var numIndex = dto.NumeroInicio + i;
            var numStr = string.IsNullOrWhiteSpace(dto.Prefijo) ? numIndex.ToString() : $"{dto.Prefijo}{numIndex}";
            
            // Si ya existe en la residencia, omitir para no romper indice único
            var existe = await db.Habitaciones.AnyAsync(h => h.ResidenciaId == dto.ResidenciaId && h.Numero == numStr);
            if (existe) continue;

            var entity = new Habitacion
            {
                ResidenciaId = dto.ResidenciaId,
                TipoHabitacionId = dto.TipoHabitacionId,
                Numero = numStr,
                Nombre = $"{numStr}",
                TipoCamaPrincipal = string.IsNullOrWhiteSpace(dto.TipoCamaPrincipal) ? "IND" : dto.TipoCamaPrincipal,
                CapacidadPersonas = dto.CapacidadPersonas,
                AdmiteSupletorias = dto.AdmiteSupletorias,
                PlazasSupletorias = dto.PlazasSupletorias,
                Activa = true,
                Orden = numIndex
            };
            db.Habitaciones.Add(entity);
            creadas.Add(entity);
        }
        await db.SaveChangesAsync();
        return await GetAllAsync(dto.ResidenciaId);
    }

    public async Task<HabitacionDto?> UpdateAsync(Guid id, UpsertHabitacionDto dto)
    {
        var entity = await db.Habitaciones.FindAsync(id);
        if (entity is null) return null;
        entity.ResidenciaId = dto.ResidenciaId; entity.TipoHabitacionId = dto.TipoHabitacionId;
        entity.Numero = dto.Numero; entity.Nombre = dto.Nombre;
        entity.TipoCamaPrincipal = string.IsNullOrWhiteSpace(dto.TipoCamaPrincipal) ? "IND" : dto.TipoCamaPrincipal;
        entity.CapacidadPersonas = dto.CapacidadPersonas;
        entity.AdmiteSupletorias = dto.AdmiteSupletorias; entity.PlazasSupletorias = dto.PlazasSupletorias;
        entity.Activa = dto.Activa; entity.Notas = dto.Notas; entity.Orden = dto.Orden;
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Habitaciones.FindAsync(id);
        if (entity is null) return false;

        var reservasAsociadas = await db.Reservas.Where(r => r.HabitacionId == id).ToListAsync();
        foreach (var r in reservasAsociadas)
        {
            r.HabitacionId = null;
        }

        db.Habitaciones.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// TARIFAS
// ─────────────────────────────────────────────────────────────────────────────
public interface ITarifaService
{
    Task<List<TarifaDto>> GetAllAsync(Guid? residenciaId = null);
    Task<TarifaDto?> GetByIdAsync(Guid id);
    Task<TarifaDto> CreateAsync(UpsertTarifaDto dto);
    Task<TarifaDto?> UpdateAsync(Guid id, UpsertTarifaDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class TarifaService(IResidenciaDbContext db) : ITarifaService
{
    public async Task<List<TarifaDto>> GetAllAsync(Guid? residenciaId = null)
    {
        var q = db.Tarifas.Include(t => t.TipoHabitacion).Include(t => t.Residencia).AsQueryable();
        if (residenciaId.HasValue)
        {
            q = q.Where(t => t.ResidenciaId == residenciaId.Value);
        }
        return await q
            .OrderBy(t => t.Residencia.Orden).ThenBy(t => t.TipoHabitacion.Orden).ThenBy(t => t.NombreTarifa)
            .Select(t => new TarifaDto(t.Id, t.ResidenciaId, t.Residencia.Nombre, t.TipoHabitacionId, t.TipoHabitacion.Nombre,
                t.NombreTarifa, t.PrecioNoche, t.PrecioMes, t.PorcentajeIva,
                t.Descripcion, t.Activa, t.VigenteDesde, t.VigenteHasta))
            .ToListAsync();
    }

    public async Task<TarifaDto?> GetByIdAsync(Guid id) =>
        await db.Tarifas.Include(t => t.TipoHabitacion).Include(t => t.Residencia).Where(t => t.Id == id)
            .Select(t => new TarifaDto(t.Id, t.ResidenciaId, t.Residencia.Nombre, t.TipoHabitacionId, t.TipoHabitacion.Nombre,
                t.NombreTarifa, t.PrecioNoche, t.PrecioMes, t.PorcentajeIva,
                t.Descripcion, t.Activa, t.VigenteDesde, t.VigenteHasta))
            .FirstOrDefaultAsync();

    public async Task<TarifaDto> CreateAsync(UpsertTarifaDto dto)
    {
        var entity = new Tarifa
        {
            ResidenciaId = dto.ResidenciaId,
            TipoHabitacionId = dto.TipoHabitacionId, NombreTarifa = dto.NombreTarifa,
            PrecioNoche = dto.PrecioNoche, PrecioMes = dto.PrecioMes,
            PorcentajeIva = dto.PorcentajeIva, Descripcion = dto.Descripcion,
            Activa = dto.Activa, VigenteDesde = dto.VigenteDesde, VigenteHasta = dto.VigenteHasta
        };
        db.Tarifas.Add(entity);
        await db.SaveChangesAsync();
        return (await GetByIdAsync(entity.Id))!;
    }

    public async Task<TarifaDto?> UpdateAsync(Guid id, UpsertTarifaDto dto)
    {
        var entity = await db.Tarifas.FindAsync(id);
        if (entity is null) return null;
        entity.ResidenciaId = dto.ResidenciaId;
        entity.TipoHabitacionId = dto.TipoHabitacionId; entity.NombreTarifa = dto.NombreTarifa;
        entity.PrecioNoche = dto.PrecioNoche; entity.PrecioMes = dto.PrecioMes;
        entity.PorcentajeIva = dto.PorcentajeIva; entity.Descripcion = dto.Descripcion;
        entity.Activa = dto.Activa; entity.VigenteDesde = dto.VigenteDesde;
        entity.VigenteHasta = dto.VigenteHasta; entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        return await GetByIdAsync(id);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Tarifas.FindAsync(id);
        if (entity is null) return false;
        db.Tarifas.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}

// ─────────────────────────────────────────────────────────────────────────────
// FESTIVOS
// ─────────────────────────────────────────────────────────────────────────────
public interface IFestivoService
{
    Task<List<FestivoDto>> GetAllAsync(int? anio = null);
    Task<FestivoDto?> GetByIdAsync(Guid id);
    Task<FestivoDto> CreateAsync(UpsertFestivoDto dto);
    Task<FestivoDto?> UpdateAsync(Guid id, UpsertFestivoDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class FestivoService(IResidenciaDbContext db) : IFestivoService
{
    public async Task<List<FestivoDto>> GetAllAsync(int? anio = null)
    {
        var query = db.Festivos.Where(f => f.Activo).AsQueryable();
        if (anio.HasValue) query = query.Where(f => f.Fecha.Year == anio.Value);
        return await query.OrderBy(f => f.Fecha)
            .Select(f => new FestivoDto(f.Id, f.Fecha, f.Descripcion, f.Ambito, f.Activo))
            .ToListAsync();
    }

    public async Task<FestivoDto?> GetByIdAsync(Guid id) =>
        await db.Festivos.Where(f => f.Id == id)
            .Select(f => new FestivoDto(f.Id, f.Fecha, f.Descripcion, f.Ambito, f.Activo))
            .FirstOrDefaultAsync();

    public async Task<FestivoDto> CreateAsync(UpsertFestivoDto dto)
    {
        var entity = new Festivo
        {
            Fecha = dto.Fecha, Descripcion = dto.Descripcion,
            Ambito = dto.Ambito, Activo = dto.Activo
        };
        db.Festivos.Add(entity);
        await db.SaveChangesAsync();
        return new FestivoDto(entity.Id, entity.Fecha, entity.Descripcion, entity.Ambito, entity.Activo);
    }

    public async Task<FestivoDto?> UpdateAsync(Guid id, UpsertFestivoDto dto)
    {
        var entity = await db.Festivos.FindAsync(id);
        if (entity is null) return null;
        entity.Fecha = dto.Fecha; entity.Descripcion = dto.Descripcion;
        entity.Ambito = dto.Ambito; entity.Activo = dto.Activo;
        await db.SaveChangesAsync();
        return new FestivoDto(entity.Id, entity.Fecha, entity.Descripcion, entity.Ambito, entity.Activo);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Festivos.FindAsync(id);
        if (entity is null) return false;
        db.Festivos.Remove(entity);
        await db.SaveChangesAsync();
        return true;
    }
}
