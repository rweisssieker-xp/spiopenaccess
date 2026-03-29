using System.Globalization;
using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Spreadsheet;

public sealed class SpreadsheetModule : IOfficeModule
{
    private readonly decimal[] _quarterlyRevenue = [182_500m, 204_400m, 198_225m, 227_000m];

    public ModuleInfo Info { get; } = new(
        "sheet",
        "Spreadsheet",
        "Tabellen, Formeln, Szenarien und Ausdrucke.",
        "Analysis",
        ["Worksheets", "Formulas", "Charts", "What-if"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var total = _quarterlyRevenue.Sum();
        var average = _quarterlyRevenue.Average();

        var content = new[]
        {
            $"Workbook         : {workspace.Name}-Financials",
            "Sheet            : FY26 Revenue",
            $"Cells in use     : {12 * 8}",
            $"Q1..Q4           : {string.Join(" | ", _quarterlyRevenue.Select(value => value.ToString("N0", CultureInfo.InvariantCulture)))}",
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
}
