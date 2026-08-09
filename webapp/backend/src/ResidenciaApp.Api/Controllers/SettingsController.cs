using System.Security.Claims;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;

namespace ResidenciaApp.Api.Controllers;

[ApiController]
[Route("settings")]
[Authorize]
public class SettingsController : ControllerBase
{
    private readonly ISettingsService _settingsService;

    public SettingsController(ISettingsService settingsService)
    {
        _settingsService = settingsService;
    }

    [HttpGet]
    public async Task<IActionResult> GetSettings()
    {
        var userId = GetUserId();
        if (userId is null)
        {
            return Unauthorized();
        }

        var settings = await _settingsService.GetSettingsAsync(userId.Value);
        return Ok(settings);
    }

    [HttpPost]
    public async Task<IActionResult> SaveSettings([FromBody] UpdateUserSettingsRequest request)
    {
        var userId = GetUserId();
        if (userId is null)
        {
            return Unauthorized();
        }

        var settings = await _settingsService.UpsertSettingsAsync(userId.Value, request.Settings);
        return Ok(settings);
    }

    private Guid? GetUserId()
    {
        var userIdClaim = User.FindFirstValue(ClaimTypes.NameIdentifier);
        return Guid.TryParse(userIdClaim, out var userId) ? userId : null;
    }
}
