namespace SpiOpenAccess.Core;

public sealed record OfficeWorkspace(
    string Name,
    string Owner,
    DateOnly SnapshotDate,
    IReadOnlyDictionary<string, string> Preferences);
