using ResidenciaApp.Application.Dtos;

namespace ResidenciaApp.Application.Interfaces;

public interface IAuthService
{
    Task<LoginResponse> LoginAsync(LoginRequest request, CancellationToken cancellationToken = default);

    Task<MeResponse> GetMeAsync(Guid userId, CancellationToken cancellationToken = default);
}
