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
}
