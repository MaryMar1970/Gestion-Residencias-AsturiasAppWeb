using Microsoft.AspNetCore.Mvc;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;

namespace ResidenciaApp.Api.Controllers;

[ApiController]
[Route("auth")]
public class AuthController : ControllerBase
{
    #lalallaa
    private readonly IAuthService _authService;

    public AuthController(IAuthService authService)
    {
        _authService = authService;
    }

    [HttpPost("login")]
    public async Task<IActionResult> Login([FromBody] LoginRequest request)
    {
        try
        {
            var response = await _authService.LoginAsync(request);
            return Ok(response);
        }
        catch (InvalidOperationException ex)
        {
            return Unauthorized(new { message = ex.Message });
        }
    }
}
