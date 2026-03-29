using System.Reflection;
using System.Text.Json;

namespace SpiOpenAccess.Modules.Database;

public static class DatabaseCatalogLoader
{
    private const string ResourceName = "SpiOpenAccess.Modules.Database.Data.openaccess-catalog.json";

    public static DatabaseCatalog LoadDefault()
    {
        using var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(ResourceName);
        if (stream is null)
        {
            throw new InvalidOperationException($"Embedded catalog resource '{ResourceName}' was not found.");
        }

        return Load(stream);
    }

    public static DatabaseCatalog Load(Stream stream)
    {
        var catalog = JsonSerializer.Deserialize<DatabaseCatalog>(stream, new JsonSerializerOptions
        {
            PropertyNameCaseInsensitive = true
        });

        if (catalog is null)
        {
            throw new InvalidOperationException("Database catalog could not be deserialized.");
        }

        return catalog;
    }
}
