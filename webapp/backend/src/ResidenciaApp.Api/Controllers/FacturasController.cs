using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Services;

namespace ResidenciaApp.Api.Controllers;

[ApiController]
[Route("api/facturas")]
[Authorize]
public class FacturasController(IFacturaService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetAll(
        [FromQuery] Guid? residenciaId = null,
        [FromQuery] int? ejercicio = null,
        [FromQuery] string? estado = null)
        => Ok(await svc.GetAllAsync(residenciaId, ejercicio, estado));

    [HttpGet("{id:guid}")]
    public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpGet("numero/{numero}")]
    public async Task<IActionResult> GetByNumero(string numero)
    {
        var r = await svc.GetByNumeroAsync(numero);
        return r is null ? NotFound() : Ok(r);
    }

    // Crear desde una reserva existente (más común)
    [HttpPost("desde-reserva/{reservaId:guid}")]
    public async Task<IActionResult> CrearDesdeReserva(Guid reservaId, [FromBody] EmitirFacturaDto? dto = null)
    {
        try
        {
            var f = await svc.CrearDesdeReservaAsync(reservaId, dto?.FormaPago);
            return CreatedAtAction(nameof(GetById), new { id = f.Id }, f);
        }
        catch (InvalidOperationException ex) { return Conflict(new { message = ex.Message }); }
    }

    // Crear manualmente con líneas libres
    [HttpPost]
    public async Task<IActionResult> CrearManual(CrearFacturaDto dto)
    {
        try
        {
            var f = await svc.CrearManualAsync(dto);
            return CreatedAtAction(nameof(GetById), new { id = f.Id }, f);
        }
        catch (Exception ex) { return BadRequest(new { message = ex.Message }); }
    }

    [HttpPost("{id:guid}/emitir")]
    public async Task<IActionResult> Emitir(Guid id, EmitirFacturaDto dto)
    {
        var r = await svc.EmitirAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpPost("{id:guid}/pagar")]
    public async Task<IActionResult> Pagar(Guid id, [FromQuery] string? formaPago = null)
    {
        var r = await svc.MarcarPagadaAsync(id, formaPago);
        return r is null ? NotFound() : Ok(r);
    }

    [HttpPost("{id:guid}/anular")]
    public async Task<IActionResult> Anular(Guid id, AnularFacturaDto dto)
    {
        var r = await svc.AnularAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }

    // Devuelve el HTML listo para imprimir (se abre en nueva pestaña)
    [HttpGet("{id:guid}/html")]
    public async Task<IActionResult> ObtenerHtml(Guid id)
    {
        try
        {
            var html = await svc.GenerarHtmlAsync(id);
            return Content(html, "text/html", System.Text.Encoding.UTF8);
        }
        catch (InvalidOperationException ex) { return NotFound(new { message = ex.Message }); }
    }
}
