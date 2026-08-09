using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Api.Controllers;

[ApiController]
[Route("api/direcciones")]
[AllowAnonymous]
public class DireccionesController(IResidenciaDbContext db) : ControllerBase
{
    [HttpGet("buscar")]
    public async Task<IActionResult> Buscar([FromQuery] string? cp = null, [FromQuery] string? municipio = null)
    {
        var q = db.CodigosPostales.AsQueryable();

        if (!string.IsNullOrWhiteSpace(cp))
        {
            var cleanCp = cp.Trim();
            q = q.Where(x => x.CodigoPostal.StartsWith(cleanCp));
        }
        else if (!string.IsNullOrWhiteSpace(municipio))
        {
            var cleanM = municipio.Trim().ToLower();
            q = q.Where(x => x.Municipio.ToLower().Contains(cleanM));
        }
        else
        {
            return BadRequest("Debe proporcionar un cp o municipio para la búsqueda.");
        }

        // Return up to 50 results to prevent large payloads
        var results = await q.OrderBy(x => x.Municipio).Take(50).ToListAsync();
        return Ok(results);
    }
}
