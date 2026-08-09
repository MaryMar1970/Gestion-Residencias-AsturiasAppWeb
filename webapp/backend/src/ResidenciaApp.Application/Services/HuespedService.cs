using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public interface IHuespedService
{
    Task<PaginatedResult<HuespedDto>> GetAllAsync(string? buscar = null, int page = 1, int pageSize = 50);
    Task<HuespedDto?> GetByIdAsync(Guid id);
    Task<HuespedDto?> GetByDniAsync(string dni);
    Task<HuespedDto> CreateAsync(UpsertHuespedDto dto);
    Task<HuespedDto?> UpdateAsync(Guid id, UpsertHuespedDto dto);
    Task<bool> DeleteAsync(Guid id);
}

public class HuespedService(IResidenciaDbContext db, ILogger<HuespedService> logger) : IHuespedService
{
    private static HuespedDto ToDto(Huesped h, int reservas = 0) => new(
        h.Id, h.Dni, h.Nombre, h.Apellidos,
        $"{h.Nombre} {h.Apellidos}",
        h.Telefono, h.Email,
        h.Direccion, h.CodigoPostal, h.Municipio, h.Provincia,
        h.CentroOrigen, h.Departamento,
        h.TipoHuesped, h.EnListaNegra, h.MotivoListaNegra,
        h.Notas, h.Empleo, h.Situacion, h.Finalidad, h.EmpleoCategoria,
        h.FamiliaNumerosa ?? "NO", h.PorcentajeDescuento,
        reservas, h.CreadoEn);

    public async Task<PaginatedResult<HuespedDto>> GetAllAsync(string? buscar = null, int page = 1, int pageSize = 50)
    {
        page = Math.Max(1, page);
        pageSize = Math.Clamp(pageSize, 1, 500);

        var q = db.Huespedes.AsQueryable();
        if (!string.IsNullOrWhiteSpace(buscar))
        {
            var b = buscar.Trim().ToLower();
            var bClean = b.Replace("-", "").Replace(" ", "").Replace(".", "");
            var tokens = b.Split(new[] { ' ', '-', '.' }, StringSplitOptions.RemoveEmptyEntries);

            q = q.Where(h =>
                h.Dni.ToLower().Contains(b) ||
                h.Dni.ToLower().Replace("-", "").Replace(" ", "").Replace(".", "").Contains(bClean) ||
                h.Nombre.ToLower().Contains(b) ||
                h.Apellidos.ToLower().Contains(b) ||
                (h.Nombre.ToLower() + " " + h.Apellidos.ToLower()).Contains(b) ||
                (h.Apellidos.ToLower() + " " + h.Nombre.ToLower()).Contains(b) ||
                (tokens.Length > 1 && tokens.All(t =>
                    h.Nombre.ToLower().Contains(t) ||
                    h.Apellidos.ToLower().Contains(t) ||
                    h.Dni.ToLower().Contains(t))) ||
                (h.Email != null && h.Email.ToLower().Contains(b)));
        }

        var totalCount = await q.CountAsync();
        var items = await q.OrderBy(h => h.Apellidos).ThenBy(h => h.Nombre)
            .Skip((page - 1) * pageSize)
            .Take(pageSize)
            .Select(h => ToDto(h, h.Reservas.Count))
            .ToListAsync();

        int totalPages = (int)Math.Ceiling((double)totalCount / pageSize);
        return new PaginatedResult<HuespedDto>(items, totalCount, page, pageSize, totalPages);
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
        logger.LogInformation("Creando nuevo huésped con DNI {Dni} - {Nombre} {Apellidos}", dto.Dni, dto.Nombre, dto.Apellidos);
        var entity = new Huesped
        {
            Dni = dto.Dni.Trim().ToUpper(), Nombre = dto.Nombre.Trim(),
            Apellidos = dto.Apellidos.Trim(), Telefono = dto.Telefono,
            Email = dto.Email, Direccion = dto.Direccion, CodigoPostal = dto.CodigoPostal,
            Municipio = dto.Municipio, Provincia = dto.Provincia,
            CentroOrigen = dto.CentroOrigen, Departamento = dto.Departamento,
            TipoHuesped = dto.TipoHuesped ?? "Externo", EnListaNegra = dto.EnListaNegra,
            MotivoListaNegra = dto.MotivoListaNegra, Notas = dto.Notas,
            Empleo = dto.Empleo, Situacion = dto.Situacion,
            Finalidad = dto.Finalidad, EmpleoCategoria = dto.EmpleoCategoria,
            FamiliaNumerosa = dto.FamiliaNumerosa ?? "NO",
            PorcentajeDescuento = dto.PorcentajeDescuento ?? 0
        };
        db.Huespedes.Add(entity);
        await db.SaveChangesAsync();
        logger.LogInformation("Huésped {HuespedId} ({Dni}) creado correctamente", entity.Id, entity.Dni);
        return ToDto(entity);
    }

    public async Task<HuespedDto?> UpdateAsync(Guid id, UpsertHuespedDto dto)
    {
        var entity = await db.Huespedes.FindAsync(id);
        if (entity is null)
        {
            logger.LogWarning("Intento de actualizar huésped inexistente {HuespedId}", id);
            return null;
        }

        logger.LogInformation("Actualizando huésped {HuespedId} ({Dni})", id, entity.Dni);
        entity.Dni = dto.Dni.Trim().ToUpper(); entity.Nombre = dto.Nombre.Trim();
        entity.Apellidos = dto.Apellidos.Trim(); entity.Telefono = dto.Telefono;
        entity.Email = dto.Email; entity.Direccion = dto.Direccion; entity.CodigoPostal = dto.CodigoPostal;
        entity.Municipio = dto.Municipio; entity.Provincia = dto.Provincia;
        entity.CentroOrigen = dto.CentroOrigen; entity.Departamento = dto.Departamento;
        if (!string.IsNullOrWhiteSpace(dto.TipoHuesped)) entity.TipoHuesped = dto.TipoHuesped;
        entity.EnListaNegra = dto.EnListaNegra;
        entity.MotivoListaNegra = dto.MotivoListaNegra; entity.Notas = dto.Notas;
        entity.Empleo = dto.Empleo; entity.Situacion = dto.Situacion;
        entity.Finalidad = dto.Finalidad; entity.EmpleoCategoria = dto.EmpleoCategoria;
        entity.FamiliaNumerosa = dto.FamiliaNumerosa ?? "NO";
        entity.PorcentajeDescuento = dto.PorcentajeDescuento ?? 0;
        entity.ActualizadoEn = DateTime.UtcNow;
        await db.SaveChangesAsync();
        logger.LogInformation("Huésped {HuespedId} actualizado correctamente", id);
        return ToDto(entity);
    }

    public async Task<bool> DeleteAsync(Guid id)
    {
        var entity = await db.Huespedes.FindAsync(id);
        if (entity is null)
        {
            logger.LogWarning("Intento de eliminar huésped inexistente {HuespedId}", id);
            return false;
        }

        logger.LogInformation("Eliminando huésped {HuespedId} ({Dni})", id, entity.Dni);
        db.Huespedes.Remove(entity);
        await db.SaveChangesAsync();
        logger.LogInformation("Huésped {HuespedId} eliminado correctamente", id);
        return true;
    }
}
