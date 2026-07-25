using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Services;

namespace ResidenciaApp.Api.Controllers;

[ApiController]
[Route("api/residencias")]
[Authorize]
public class ResidenciasController(IResidenciaService svc) : ControllerBase
{
    [HttpGet] public async Task<IActionResult> GetAll() => Ok(await svc.GetAllAsync());
    [HttpGet("{id:guid}")] public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpPost][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create(UpsertResidenciaDto dto)
    {
        var r = await svc.CreateAsync(dto);
        return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
    }
    [HttpPut("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, UpsertResidenciaDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpDelete("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

[ApiController]
[Route("api/tipos-habitacion")]
[Authorize]
public class TiposHabitacionController(ITipoHabitacionService svc) : ControllerBase
{
    [HttpGet] public async Task<IActionResult> GetAll() => Ok(await svc.GetAllAsync());
    [HttpGet("{id:guid}")] public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpPost][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create(UpsertTipoHabitacionDto dto)
    {
        var r = await svc.CreateAsync(dto);
        return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
    }
    [HttpPut("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, UpsertTipoHabitacionDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpDelete("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

[ApiController]
[Route("api/habitaciones")]
[Authorize]
public class HabitacionesController(IHabitacionService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetAll([FromQuery] Guid? residenciaId = null)
        => Ok(await svc.GetAllAsync(residenciaId));
    [HttpGet("{id:guid}")] public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpPost][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create(UpsertHabitacionDto dto)
    {
        var r = await svc.CreateAsync(dto);
        return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
    }
    [HttpPut("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, UpsertHabitacionDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpDelete("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

[ApiController]
[Route("api/tarifas")]
[Authorize]
public class TarifasController(ITarifaService svc) : ControllerBase
{
    [HttpGet] public async Task<IActionResult> GetAll([FromQuery] Guid? residenciaId = null) => Ok(await svc.GetAllAsync(residenciaId));
    [HttpGet("{id:guid}")] public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpPost][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create(UpsertTarifaDto dto)
    {
        var r = await svc.CreateAsync(dto);
        return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
    }
    [HttpPut("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, UpsertTarifaDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpDelete("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}

[ApiController]
[Route("api/festivos")]
[Authorize]
public class FestivosController(IFestivoService svc) : ControllerBase
{
    [HttpGet]
    public async Task<IActionResult> GetAll([FromQuery] int? anio = null)
        => Ok(await svc.GetAllAsync(anio));
    [HttpGet("{id:guid}")] public async Task<IActionResult> GetById(Guid id)
    {
        var r = await svc.GetByIdAsync(id);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpPost][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Create(UpsertFestivoDto dto)
    {
        var r = await svc.CreateAsync(dto);
        return CreatedAtAction(nameof(GetById), new { id = r.Id }, r);
    }
    [HttpPut("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Update(Guid id, UpsertFestivoDto dto)
    {
        var r = await svc.UpdateAsync(id, dto);
        return r is null ? NotFound() : Ok(r);
    }
    [HttpDelete("{id:guid}")][Authorize(Roles = "Admin")]
    public async Task<IActionResult> Delete(Guid id)
        => await svc.DeleteAsync(id) ? NoContent() : NotFound();
}
