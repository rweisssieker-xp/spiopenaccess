using SpiOpenAccess.App;

namespace SpiOpenAccess.Tests;

public sealed class AppSessionStoreTests
{
    [Fact]
    public void SaveAndLoad_RoundTripsWorkspaceState()
    {
        var root = Path.Combine(Path.GetTempPath(), "spi-retro-tests", Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        var store = new AppSessionStore(root);
        var state = new AppSessionState();
        state.Spreadsheet.Q1 = 190000m;
        state.Word.Title = "Demo Draft";
        state.Mail.Subject = "Saved subject";
        state.Communications.CurrentTarget = "REMOTE";
        state.Reporting.ActiveLayout = "AgingDetail";
        state.Programming.ProgramName = "custom.pro";

        store.Save(state);
        var reloaded = store.Load();

        Assert.Equal(190000m, reloaded.Spreadsheet.Q1);
        Assert.Equal("Demo Draft", reloaded.Word.Title);
        Assert.Equal("Saved subject", reloaded.Mail.Subject);
        Assert.Equal("REMOTE", reloaded.Communications.CurrentTarget);
        Assert.Equal("AgingDetail", reloaded.Reporting.ActiveLayout);
        Assert.Equal("custom.pro", reloaded.Programming.ProgramName);
    }
}
