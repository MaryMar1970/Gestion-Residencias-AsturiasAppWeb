using Microsoft.EntityFrameworkCore;
using ResidenciaApp.Application.Dtos;
using ResidenciaApp.Application.Interfaces;
using ResidenciaApp.Domain.Entities;

namespace ResidenciaApp.Application.Services;

public class SettingsService : ISettingsService
{
    private readonly IResidenciaDbContext _context;

    public SettingsService(IResidenciaDbContext context)
    {
        _context = context;
    }

    public async Task<SettingsResponse> GetSettingsAsync(Guid userId, CancellationToken cancellationToken = default)
    {
        var systemSettings = await _context.SystemSettings
            .OrderBy(x => x.Category)
            .ThenBy(x => x.Key)
            .Select(x => new SystemSettingDto(x.Key, x.Value, x.Description, x.Category))
            .ToListAsync(cancellationToken);

        var userSettings = await _context.UserSettings
            .Where(x => x.UserId == userId)
            .OrderBy(x => x.Key)
            .Select(x => new UserSettingDto(x.Key, x.Value, x.Description))
            .ToListAsync(cancellationToken);

        return new SettingsResponse(systemSettings, userSettings);
    }

    public async Task<SettingsResponse> UpsertSettingsAsync(Guid userId, IReadOnlyList<UpsertSettingRequest> settings, CancellationToken cancellationToken = default)
    {
        foreach (var setting in settings)
        {
            var existing = await _context.UserSettings
                .FirstOrDefaultAsync(x => x.UserId == userId && x.Key == setting.Key, cancellationToken);

            if (existing is null)
            {
                _context.UserSettings.Add(new UserSetting
                {
                    UserId = userId,
                    Key = setting.Key,
                    Value = setting.Value,
                    Description = "Configuración personalizada"
                });
            }
            else
            {
                existing.Value = setting.Value;
                existing.UpdatedAt = DateTime.UtcNow;
            }
        }

        await _context.SaveChangesAsync(cancellationToken);
        return await GetSettingsAsync(userId, cancellationToken);
    }
}
