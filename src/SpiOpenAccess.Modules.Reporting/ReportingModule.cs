using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Database;

namespace SpiOpenAccess.Modules.Reporting;

public sealed class ReportingModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "report",
        "Reporting",
        "Listings, grouping, batch output, and print preparation.",
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

    public ModuleScreen BuildRunScreen(string reportName)
    {
        return ModuleScreen.Create(
            $"Run {reportName}",
            "Batch execution for a report definition.",
            new[]
            {
                $"Definition       : {reportName}",
                "Input source     : Shared workspace",
                "Rows scanned     : 3,412",
                "Groups           : 12",
                "Output target    : PDF spool + Laser A",
                "Status           : Completed"
            },
            ["schedule Evening Finance", "design Revenue-City", "back"]);
    }

    public ModuleScreen BuildRunScreen(string reportName, DatabaseWorkspaceState databaseState)
    {
        var table = databaseState.GetTable("INVOICES");
        var groups = table.Records
            .GroupBy(record => record.GetValueOrDefault("Status", "<null>"))
            .OrderBy(group => group.Key)
            .ToArray();

        var content = new List<string>
        {
            $"Definition       : {reportName}",
            "Input source     : Session database state",
            $"Rows scanned     : {table.Records.Count}",
            $"Groups           : {groups.Length}",
            "Preview          :"
        };
        content.AddRange(groups.Select(group => $"  {group.Key,-12} count={group.Count()}"));
        content.Add("Output target    : PDF spool + Laser A");
        content.Add("Status           : Completed");

        return ModuleScreen.Create(
            $"Run {reportName}",
            "Batch execution for a report definition.",
            content,
            ["schedule Evening Finance", "design Revenue-City", "back"]);
    }

    public ModuleScreen BuildScheduleScreen(string queueName)
    {
        return ModuleScreen.Create(
            "Batch Schedule",
            "Schedule recurring report runs.",
            new[]
            {
                $"Queue            : {queueName}",
                "Start time       : 18:30",
                "Recurrence       : Weekdays",
                "Printer target   : Laser A",
                "Retention        : 14 days"
            },
            ["run Aging", "design Revenue-City", "back"]);
    }

    public ModuleScreen BuildDesignScreen(string reportName)
    {
        return ModuleScreen.Create(
            "Report Designer",
            "Layout definition for reports and listings.",
            new[]
            {
                $"Layout           : {reportName}",
                "Sections         : Header, Detail, Footer",
                "Grouping         : Enabled",
                "Summary fields   : Count, NetAmount",
                "Printer escapes  : Dot matrix compatible"
            },
            ["run Aging", "schedule Evening Finance", "back"]);
    }
}
