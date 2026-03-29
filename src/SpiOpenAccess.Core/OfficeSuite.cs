namespace SpiOpenAccess.Core;

public sealed class OfficeSuite
{
    private readonly IReadOnlyList<IOfficeModule> _modules;

    public OfficeSuite(string name, string version, OfficeWorkspace workspace, IEnumerable<IOfficeModule> modules)
    {
        Name = name;
        Version = version;
        Workspace = workspace;
        _modules = modules.OrderBy(module => module.Info.DisplayName, StringComparer.OrdinalIgnoreCase).ToArray();
    }

    public string Name { get; }

    public string Version { get; }

    public OfficeWorkspace Workspace { get; }

    public IReadOnlyList<IOfficeModule> Modules => _modules;

    public IOfficeModule? FindModule(string selector)
    {
        return _modules.FirstOrDefault(module =>
            string.Equals(module.Info.Id, selector, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(module.Info.DisplayName, selector, StringComparison.OrdinalIgnoreCase));
    }
}
