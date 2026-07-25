namespace ResidenciaApp.Application.Dtos;

public record LoginRequest(string Email, string Password);

public record LoginResponse(string Token, string Email, string FullName);

public record MeResponse(Guid Id, string Email, string FullName, IReadOnlyList<string> Roles);

public record SettingsResponse(IReadOnlyList<SystemSettingDto> SystemSettings, IReadOnlyList<UserSettingDto> UserSettings);

public record SystemSettingDto(string Key, string Value, string? Description, string Category);

public record UserSettingDto(string Key, string Value, string? Description);

public record UpsertSettingRequest(string Key, string Value);

public record UpdateUserSettingsRequest(IReadOnlyList<UpsertSettingRequest> Settings);
