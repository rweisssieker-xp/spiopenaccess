using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Reporting;

public sealed class ReportingModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "report",
        "Reporting",
        "Listen, Gruppierungen, Summenlaeufe und Druckaufbereitung.",
        "Output",
        ["Layouts", "Grouping", "Batch Run", "Printer Targets"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var content = new[]
        {
            $"Catalog          : {workspace.Name} Reporting",
            "Definitions      : 18",
            "Batch queues     : Morning Ops, Evening Finance",
            "Printer targets  : Laser A, Dot Matrix, PDF spool",
            "Recent output    : Aging-2026-03-28.lst, Revenue-City.prn",
            "Layout blocks    : Header, Detail, Group Footer, Summary"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["run Aging", "schedule Evening Finance", "design Revenue-City"]);
    }
}
