using System.Globalization;
using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Spreadsheet;

public sealed class SpreadsheetModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "sheet",
        "Spreadsheet",
        "Worksheets, formulas, scenarios, and print previews.",
        "Analysis",
        ["Worksheets", "Formulas", "Charts", "What-if"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        return BuildHomeScreen(workspace, new SpreadsheetWorkspaceState());
    }

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace, SpreadsheetWorkspaceState state)
    {
        var revenue = state.ToArray();
        var total = revenue.Sum();
        var average = revenue.Average();

        var content = new[]
        {
            $"Workbook         : {workspace.Name}-Financials",
            "Sheet            : FY26 Revenue",
            $"Cells in use     : {12 * 8}",
            $"Q1..Q4           : {string.Join(" | ", revenue.Select(value => value.ToString("N0", CultureInfo.InvariantCulture)))}",
            $"Total            : {total.ToString("N0", CultureInfo.InvariantCulture)}",
            $"Average quarter  : {average.ToString("N2", CultureInfo.InvariantCulture)}",
            "Named ranges     : RevenueQ1, RevenueQ2, RevenueQ3, RevenueQ4",
            "Scenario manager : Conservative, Base, Aggressive"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["recalc", "goal-seek margin 18", "print area A1:H56"]);
    }

    public ModuleScreen BuildRecalcScreen(OfficeWorkspace workspace, SpreadsheetWorkspaceState state)
    {
        var revenue = state.ToArray();
        var total = revenue.Sum();
        var margins = revenue.Select(value => Math.Round(value / total * 100m, 2)).ToArray();

        return ModuleScreen.Create(
            "Recalc Worksheet",
            "Recalculate the active workbook.",
            new[]
            {
                $"Workbook         : {workspace.Name}-Financials",
                "Sheet            : FY26 Revenue",
                $"Recalc timestamp : {workspace.SnapshotDate:yyyy-MM-dd} 09:00",
                $"Total revenue    : {total.ToString("N0", CultureInfo.InvariantCulture)}",
                $"Quarter mix      : Q1={margins[0]}%  Q2={margins[1]}%  Q3={margins[2]}%  Q4={margins[3]}%",
                "Dependency graph : 12 formulas refreshed",
                "Status           : Workbook consistent"
            },
            ["goal-seek margin 18", "print area A1:H56", "back"]);
    }

    public ModuleScreen BuildGoalSeekScreen(string target)
    {
        return ModuleScreen.Create(
            "Goal Seek",
            "Scenario analysis for target values.",
            new[]
            {
                "Objective cell    : Margin",
                $"Requested target  : {target}",
                "Driver cell       : RevenueQ4",
                "Iterations        : 6",
                "Suggested value   : 239,400",
                "Result            : Feasible in aggressive scenario"
            },
            ["recalc", "print area A1:H56", "back"]);
    }

    public ModuleScreen BuildPrintAreaScreen(string area)
    {
        return ModuleScreen.Create(
            "Print Area",
            "Print preview for the selected cell range.",
            new[]
            {
                $"Range            : {area}",
                "Orientation      : Landscape",
                "Scaling          : Fit to 1 page wide",
                "Headers          : Enabled",
                "Footer           : Page x of y",
                "Preview          : FY26 Revenue overview"
            },
            ["recalc", "goal-seek margin 18", "back"]);
    }
}
