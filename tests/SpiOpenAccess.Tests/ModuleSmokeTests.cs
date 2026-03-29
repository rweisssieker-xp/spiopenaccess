using SpiOpenAccess.Infrastructure;

namespace SpiOpenAccess.Tests;

public sealed class ModuleSmokeTests
{
    [Fact]
    public void EveryModule_BuildsNonEmptyHomeScreen()
    {
        var suite = OfficeSuiteFactory.CreateDefault();

        foreach (var module in suite.Modules)
        {
            var screen = module.BuildHomeScreen(suite.Workspace);

            Assert.False(string.IsNullOrWhiteSpace(screen.Title));
            Assert.False(string.IsNullOrWhiteSpace(screen.Summary));
            Assert.NotEmpty(screen.Content);
            Assert.NotEmpty(screen.Commands);
        }
    }

    [Fact]
    public void EveryModule_HasAtLeastThreeCapabilities()
    {
        var suite = OfficeSuiteFactory.CreateDefault();

        foreach (var module in suite.Modules)
        {
            Assert.True(module.Info.Capabilities.Length >= 3, $"Module {module.Info.Id} should expose multiple capabilities.");
        }
    }
}
