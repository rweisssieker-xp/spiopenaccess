namespace SpiOpenAccess.Core;

public sealed record ModuleInfo(
    string Id,
    string DisplayName,
    string Summary,
    string Category,
    string[] Capabilities);
