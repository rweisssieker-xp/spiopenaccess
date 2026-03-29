using SpiOpenAccess.Infrastructure;
using SpiOpenAccess.Modules.Database;

namespace SpiOpenAccess.Tests;

public sealed class OfficeSuiteFactoryTests
{
    [Fact]
    public void CreateDefault_RegistersExpectedModules()
    {
        var suite = OfficeSuiteFactory.CreateDefault();

        var ids = suite.Modules.Select(module => module.Info.Id).ToArray();

        Assert.Equal(
            ["comm", "db", "mail", "pro", "report", "sheet", "word"],
            ids);
    }

    [Fact]
    public void CreateDefault_WiresDatabaseModuleWithSeedCatalog()
    {
        var suite = OfficeSuiteFactory.CreateDefault();

        var databaseModule = Assert.IsType<DatabaseModule>(suite.FindModule("db"));
        var customers = databaseModule.FindTable("CUSTOMERS");

        Assert.NotNull(customers);
        Assert.Equal(3, customers.Records.Count);
    }
}
