using ResidenciaApp.Application.Dtos;

namespace ResidenciaApp.Application.Interfaces;

public interface ISettingsService
{
    Task<SettingsResponse> GetSettingsAsync(Guid userId, CancellationToken cancellationToken = default);

    Task<SettingsResponse> UpsertSettingsAsync(Guid userId, IReadOnlyList<UpsertSettingRequest> settings, CancellationToken cancellationToken = default);
}
