using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Spreadsheet;

namespace SpiOpenAccess.Tests;

public sealed class SpreadsheetModuleTests
{
    [Fact]
    public void BuildHomeScreen_ContainsCalculatedTotals()
    {
        var module = new SpreadsheetModule();
        var workspace = new OfficeWorkspace(
            "OPENACCESS",
            "ADMIN",
            new DateOnly(2026, 3, 28),
            new Dictionary<string, string>());

        var screen = module.BuildHomeScreen(workspace);

        Assert.Contains(screen.Content, line => line.Contains("812,125", StringComparison.Ordinal));
    }

    [Fact]
    public void GridEditingAndSum_AreComputedFromState()
    {
        var module = new SpreadsheetModule();
        var state = new SpreadsheetWorkspaceState();
        state.SelectCell("B2");
        state.SetCell("B2", 250000m);

        var grid = module.BuildGridScreen(state);
        var sum = module.BuildCellSumScreen(state, "A2", "D2");

        Assert.Contains(grid.Content, line => line.Contains("Active cell      : B2", StringComparison.Ordinal));
        Assert.Contains(sum.Content, line => line.Contains("Result", StringComparison.Ordinal));
    }
}
