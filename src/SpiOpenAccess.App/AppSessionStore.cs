using System.Text.Json;

namespace SpiOpenAccess.App;

public sealed class AppSessionStore
{
    private readonly string _statePath;

    public AppSessionStore(string rootDirectory)
    {
        _statePath = Path.Combine(rootDirectory, ".retro-state", "session.json");
    }

    public AppSessionState Load()
    {
        if (!File.Exists(_statePath))
        {
            return new AppSessionState();
        }

        using var stream = File.OpenRead(_statePath);
        return JsonSerializer.Deserialize<AppSessionState>(stream, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        }) ?? new AppSessionState();
    }

    public void Save(AppSessionState state)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(_statePath)!);
        using var stream = File.Create(_statePath);
        JsonSerializer.Serialize(stream, state, new JsonSerializerOptions
        {
            WriteIndented = true
        });
    }
}
