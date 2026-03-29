namespace SpiOpenAccess.Core;

public interface IOfficeModule
{
    ModuleInfo Info { get; }

    ModuleScreen BuildHomeScreen(OfficeWorkspace workspace);
}
