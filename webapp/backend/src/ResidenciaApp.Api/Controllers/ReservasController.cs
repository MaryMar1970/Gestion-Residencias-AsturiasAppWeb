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
    public async Task<IActionResult> GetAll([FromQuery] string? buscar = null, [FromQuery] int page = 1, [FromQuery] int pageSize = 50)
        => Ok(await svc.GetAllAsync(buscar, page, pageSize));

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
        [FromQuery] bool incluirCanceladas = false,
        [FromQuery] int page = 1,
        [FromQuery] int pageSize = 50)
        => Ok(await svc.GetAllAsync(residenciaId, habitacionId, fechaDesde, fechaHasta, incluirCanceladas, page, pageSize));

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpGet("comprobar-solapamiento")]
    public async Task<IActionResult> ComprobarSolapamiento(
        [FromQuery] Guid? habitacionId,
        [FromQuery] string fechaEntrada,
        [FromQuery] string fechaSalida,
        [FromQuery] Guid? excluirReservaId = null,
        [FromQuery] bool esBloqueo = false)
        => Ok(await svc.ComprobarSolapamientoAsync(habitacionId, fechaEntrada, fechaSalida, excluirReservaId, esBloqueo));

    [HttpGet("habitaciones-disponibles")]
    public async Task<IActionResult> GetHabitacionesDisponibles(
        [FromQuery] Guid residenciaId,
        [FromQuery] string fechaEntrada,
        [FromQuery] string fechaSalida,
        [FromQuery] int pax = 1)
        => Ok(await svc.GetHabitacionesDisponiblesAsync(residenciaId, fechaEntrada, fechaSalida, pax));

    [HttpPost]
    public async Task<IActionResult> Create(CrearReservaDto dto)
    {
        try
        {
            var res = await svc.CreateAsync(dto);
            return CreatedAtAction(nameof(GetById), new { id = res.Reserva.Id }, res);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }

    [HttpPut("{id:guid}")]
    public async Task<IActionResult> Update(Guid id, ActualizarReservaDto dto)
    {
        try
        {
            var res = await svc.UpdateAsync(id, dto);
            return res is null ? NotFound() : Ok(res);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }

    [HttpPost("{id:guid}/mover")]
    public async Task<IActionResult> Mover(Guid id, MoverReservaDto dto)
    {
        try
        {
            var res = await svc.MoverAsync(id, dto);
            return res is null ? NotFound() : Ok(res);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }



    [HttpGet("{id:guid}/candidatos-reevaluacion")]
    public async Task<IActionResult> GetCandidatosReevaluacion(Guid id)
        => Ok(await svc.GetCandidatosReevaluacionAsync(id));

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
        [FromQuery] string? fechaInicio = null,
        [FromQuery] string? fechaFin = null,
        [FromQuery] string? inicio = null,
        [FromQuery] string? fin = null,
        [FromQuery] int? dias = null,
        [FromQuery] Guid? residenciaId = null)
    {
        var startStr = fechaInicio ?? inicio ?? DateOnly.FromDateTime(DateTime.Today).ToString("yyyy-MM-dd");
        var endStr = fechaFin ?? fin;
        if (string.IsNullOrEmpty(endStr))
        {
            if (DateOnly.TryParse(startStr, out var startDate))
            {
                var numDias = dias.HasValue && dias.Value > 0 ? dias.Value : 10;
                endStr = startDate.AddDays(numDias).ToString("yyyy-MM-dd");
            }
            else
            {
                endStr = startStr;
            }
        }
        return Ok(await svc.GetCalendarioAsync(startStr, endStr, residenciaId));
    }
}
