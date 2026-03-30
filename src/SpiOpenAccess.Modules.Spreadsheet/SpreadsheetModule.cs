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
            $"Active cell      : {state.ActiveCell} = {state.GetCell(state.ActiveCell).ToString("N2", CultureInfo.InvariantCulture)}",
            "Named ranges     : RevenueQ1, RevenueQ2, RevenueQ3, RevenueQ4",
            "Scenario manager : Conservative, Base, Aggressive"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["grid", "select A1", "put A1 120000", "recalc", "goal-seek margin 18", "print area A1:H56"]);
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

    public ModuleScreen BuildGridScreen(SpreadsheetWorkspaceState state)
    {
        var rows = new List<string>
        {
            $"Active cell      : {state.ActiveCell}",
            "Grid             :",
            "        A           B           C           D"
        };

        for (var row = 1; row <= 4; row++)
        {
            rows.Add($"  {row,2}  {RenderCell(state, $"A{row}")} {RenderCell(state, $"B{row}")} {RenderCell(state, $"C{row}")} {RenderCell(state, $"D{row}")}");
        }

        rows.Add($"Cell value       : {state.GetCell(state.ActiveCell).ToString("N2", CultureInfo.InvariantCulture)}");

        return ModuleScreen.Create(
            "Worksheet Grid",
            "Editable sheet grid for the active workbook.",
            rows,
            ["select A1", "put A1 120000", "sum A1 D1", "recalc", "back"]);
    }

    public ModuleScreen BuildCellSumScreen(SpreadsheetWorkspaceState state, string fromCell, string toCell)
    {
        var from = SpreadsheetWorkspaceState.NormalizeCell(fromCell);
        var to = SpreadsheetWorkspaceState.NormalizeCell(toCell);
        if (from.Length != 2 || to.Length != 2 || from[1] != to[1])
        {
            throw new InvalidOperationException("Only horizontal ranges are supported, example: sum A1 D1");
        }

        var row = from[1..];
        var colStart = from[0];
        var colEnd = to[0];
        if (colStart > colEnd)
        {
            (colStart, colEnd) = (colEnd, colStart);
        }

        decimal total = 0m;
        for (var col = colStart; col <= colEnd; col++)
        {
            total += state.GetCell($"{col}{row}");
        }

        return ModuleScreen.Create(
            "Range Sum",
            "Computed range total from the active worksheet.",
            new[]
            {
                $"Range            : {from}:{to}",
                $"Result           : {total.ToString("N2", CultureInfo.InvariantCulture)}"
            },
            ["grid", "select A1", "put A1 120000", "back"]);
    }

    private static string RenderCell(SpreadsheetWorkspaceState state, string cell)
    {
        var value = state.GetCell(cell).ToString("N0", CultureInfo.InvariantCulture).PadLeft(8);
        return string.Equals(state.ActiveCell, cell, StringComparison.OrdinalIgnoreCase)
            ? $"[{value}]"
            : $" {value} ";
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
