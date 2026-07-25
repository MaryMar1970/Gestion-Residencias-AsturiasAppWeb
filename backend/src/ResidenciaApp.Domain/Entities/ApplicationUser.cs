using Microsoft.AspNetCore.Identity;

namespace ResidenciaApp.Domain.Entities;

public class ApplicationUser : IdentityUser<Guid>
{
    public string FullName { get; set; } = string.Empty;

    public DateTime CreatedAt { get; set; } = DateTime.UtcNow;

    public ICollection<UserSetting> Settings { get; set; } = new List<UserSetting>();
}
