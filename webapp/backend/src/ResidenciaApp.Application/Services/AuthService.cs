using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text;
using Microsoft.AspNetCore.Identity;
using Microsoft.Extensions.Options;
using Microsoft.IdentityModel.Tokens;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Application.Options;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public class AuthService : IAuthService
{
    private readonly UserManager<ApplicationUser> _userManager;
    private readonly SignInManager<ApplicationUser> _signInManager;
    private readonly JwtSettings _jwtSettings;

    public AuthService(UserManager<ApplicationUser> userManager, SignInManager<ApplicationUser> signInManager, IOptions<JwtSettings> jwtSettings)
    {
        _userManager = userManager;
        _signInManager = signInManager;
        _jwtSettings = jwtSettings.Value;
    }

    public async Task<LoginResponse> LoginAsync(LoginRequest request, CancellationToken cancellationToken = default)
    {
        var user = await _userManager.FindByEmailAsync(request.Email);
        if (user is null)
        {
            throw new InvalidOperationException("Credenciales inválidas.");
        }

        var result = await _signInManager.CheckPasswordSignInAsync(user, request.Password, false);
        if (!result.Succeeded)
        {
            throw new InvalidOperationException("Credenciales inválidas.");
        }

        var roles = (await _userManager.GetRolesAsync(user)).ToList();
        var token = CreateToken(user, roles);

        return new LoginResponse(token, user.Email ?? string.Empty, user.FullName);
    }

    public async Task<MeResponse> GetMeAsync(Guid userId, CancellationToken cancellationToken = default)
    {
        var user = await _userManager.FindByIdAsync(userId.ToString());
        if (user is null)
        {
            throw new InvalidOperationException("Usuario no encontrado.");
        }

        var roles = await _userManager.GetRolesAsync(user);
        return new MeResponse(user.Id, user.Email ?? string.Empty, user.FullName, roles.ToList());
    }

    private string CreateToken(ApplicationUser user, IReadOnlyList<string> roles)
    {
        var secretStr = string.IsNullOrWhiteSpace(_jwtSettings.Secret) || _jwtSettings.Secret.Length < 16
            ? "super-secret-key-for-local-development-123456"
            : _jwtSettings.Secret;
        var issuerStr = string.IsNullOrWhiteSpace(_jwtSettings.Issuer) ? "ResidenciaApp" : _jwtSettings.Issuer;
        var audienceStr = string.IsNullOrWhiteSpace(_jwtSettings.Audience) ? "ResidenciaAppClient" : _jwtSettings.Audience;

        var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(secretStr));
        var credentials = new SigningCredentials(key, SecurityAlgorithms.HmacSha256);

        var claims = new List<Claim>
        {
            new(ClaimTypes.NameIdentifier, user.Id.ToString()),
            new(ClaimTypes.Email, user.Email ?? string.Empty),
            new(ClaimTypes.Name, user.FullName)
        };

        claims.AddRange(roles.Select(role => new Claim(ClaimTypes.Role, role)));

        var token = new JwtSecurityToken(
            issuer: issuerStr,
            audience: audienceStr,
            claims: claims,
            expires: DateTime.UtcNow.AddMinutes(_jwtSettings.ExpirationMinutes > 0 ? _jwtSettings.ExpirationMinutes : 60),
            signingCredentials: credentials);

        return new JwtSecurityTokenHandler().WriteToken(token);
    }
}
