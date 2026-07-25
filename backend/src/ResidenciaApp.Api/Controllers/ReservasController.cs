using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Services;

namespace ResidenciaApp.Api.Controllers;

// ─────────────────────────────────────────────────────────────────────────────
// HUÉSPEDES
// ─────────────────────────────────────────────────────────────────────────────
[ApiController]
[Route("api/huespedes")]
[Authorize]
public class HuespedesController(IHuespedService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetAll([FromQuery] string? buscar = null)
        => Ok(await svc.GetAllAsync(buscar));

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpGet("dni/{dni}")]
    public async Task<IActionResult> GetByDni(string dni)
    {
        var r = await svc.GetByDniAsync(dni);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpPost]
    public async Task<IActionResult> Create(UpsertHuespedDto dto)
    {
        try
        {
            var r = await svc.CreateAsync(dto);
            return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
        }
        catch (Exception ex) { return BadRequest(new { message = ex.Message }); }
    }

    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, UpsertHuespedDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpDelete("{id:guid}")]
    [Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

// ─────────────────────────────────────────────────────────────────────────────
// RESERVAS
// ─────────────────────────────────────────────────────────────────────────────
[ApiController]
[Route("api/reservas")]
[Authorize]
public class ReservasController(IReservaService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] Guid? residenciaId = null,
        [FromQuery] Guid? habitacionId = null,
        [FromQuery] string? fechaDesde = null,
        [FromQuery] string? fechaHasta = null,
        [FromQuery] bool incluirCanceladas = false)
        => Ok(await svc.GetAllAsync(residenciaId, habitacionId, fechaDesde, fechaHasta, incluirCanceladas));

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpGet("comprobar-solapamiento")]
    public async Task<IActionResult> ComprobarSolapamiento(
        [FromQuery] Guid habitacionId,
        [FromQuery] string fechaEntrada,
        [FromQuery] string fechaSalida,
        [FromQuery] Guid? excluirReservaId = null,
        [FromQuery] bool esBloqueo = false)
        => Ok(await svc.ComprobarSolapamientoAsync(habitacionId, fechaEntrada, fechaSalida, excluirReservaId, esBloqueo));

    [HttpPost]
    public async Task<IActionResult> Create(CrearReservaDto dto)
    {
        try
        {
            var r = await svc.CreateAsync(dto);
            return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }

    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, ActualizarReservaDto dto)
    {
        try
        {
            var r = await svc.UpdateAsync(id, dto);
            return r is null ? NotFound() : Ok(r);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }

    [HttpDelete("{id:guid}")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

// ─────────────────────────────────────────────────────────────────────────────
// CALENDARIO
// ─────────────────────────────────────────────────────────────────────────────
[ApiController]
[Route("api/calendario")]
[Authorize]
public class CalendarioController(IReservaService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetCalendario(
        [FromQuery] string fechaInicio,
        [FromQuery] string fechaFin,
        [FromQuery] Guid? residenciaId = null)
        => Ok(await svc.GetCalendarioAsync(fechaInicio, fechaFin, residenciaId));
}
