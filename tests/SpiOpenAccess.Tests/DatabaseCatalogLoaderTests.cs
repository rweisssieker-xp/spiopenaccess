using System.Text;
using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Database;

namespace SpiOpenAccess.Tests;

public sealed class DatabaseCatalogLoaderTests
{
    [Fact]
    public void Load_ParsesTablesFormsAndReports()
    {
        const string json = """
            {
              "tables": [
                {
                  "name": "CUSTOMERS",
                  "keyField": "Id",
                  "columns": [
                    { "name": "Id", "type": "CHAR", "width": 8 }
                  ],
                  "records": [
                    { "Id": "C-1" }
                  ]
                }
              ],
              "forms": [
                {
                  "name": "CUSTOMER_CARD",
                  "table": "CUSTOMERS",
                  "fields": [ "Id" ]
                }
              ],
              "reports": [
                {
                  "name": "AGING",
                  "table": "CUSTOMERS",
                  "groupBy": "Id",
                  "summaries": [ "count" ]
                }
              ]
            }
            """;

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(json));
        var catalog = DatabaseCatalogLoader.Load(stream);

        Assert.Single(catalog.Tables);
        Assert.Single(catalog.Forms);
        Assert.Single(catalog.Reports);
        Assert.Equal("C-1", catalog.Tables[0].Records[0]["Id"]);
    }

    [Fact]
    public void DatabaseModule_BuildHomeScreen_ShowsCurrentRecordAndLayouts()
    {
        var module = new DatabaseModule(DatabaseCatalogLoader.LoadDefault());
        var workspace = new OfficeWorkspace(
            "OPENACCESS",
            "ADMIN",
            new DateOnly(2026, 3, 28),
            new Dictionary<string, string>());

        var screen = module.BuildHomeScreen(workspace);

        Assert.Contains(screen.Content, line => line.Contains("CUSTOMER_CARD", StringComparison.Ordinal));
        Assert.Contains(screen.Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(screen.Content, line => line.Contains("AGING", StringComparison.Ordinal));
    }

    [Fact]
    public void DatabaseModule_BuildTableFormAndReportScreens_ReturnExpectedDetails()
    {
        var module = new DatabaseModule(DatabaseCatalogLoader.LoadDefault());

        var tableScreen = module.BuildTableScreen("CUSTOMERS");
        var formScreen = module.BuildFormScreen("CUSTOMER_CARD");
        var reportScreen = module.BuildReportScreen("AGING");

        Assert.Contains(tableScreen.Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(formScreen.Content, line => line.Contains("Bound table      : CUSTOMERS", StringComparison.Ordinal));
        Assert.Contains(reportScreen.Content, line => line.Contains("OPEN", StringComparison.Ordinal));
    }
}
