namespace ResidenciaApp.Domain.Entities;

public class SystemSetting
{
    public Guid Id { get; set; } = Guid.NewGuid();

    public string Key { get; set; } = string.Empty;

    public string Value { get; set; } = string.Empty;

    public string? Description { get; set; }

    public string Category { get; set; } = "General";

    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    public DateTime UpdatedAt { get; set; } = DateTime.UtcNow;
}
