namespace SpiOpenAccess.Core;

public sealed record ModuleScreen(
    string Title,
    string Summary,
    IReadOnlyList<string> Content,
    IReadOnlyList<string> Commands)
{
    public static ModuleScreen Create(
        string title,
        string summary,
        IEnumerable<string> content,
        IEnumerable<string>? commands = null)
    {
        return new ModuleScreen(
            title,
            summary,
            content.ToArray(),
            commands?.ToArray() ?? Array.Empty<string>());
    }
}
