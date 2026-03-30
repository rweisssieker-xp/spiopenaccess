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
        var state = module.CreateWorkspaceState();

        var tableScreen = module.BuildTableScreen(state, "CUSTOMERS");
        var formScreen = module.BuildFormScreen(state, "CUSTOMER_CARD");
        var reportScreen = module.BuildReportScreen(state, "AGING");

        Assert.Contains(tableScreen.Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(formScreen.Content, line => line.Contains("Bound table      : CUSTOMERS", StringComparison.Ordinal));
        Assert.Contains(reportScreen.Content, line => line.Contains("OPEN", StringComparison.Ordinal));
    }

    [Fact]
    public void DatabaseModule_CanSearchAppendAndUpdateRecords()
    {
        var module = new DatabaseModule(DatabaseCatalogLoader.LoadDefault());
        var state = module.CreateWorkspaceState();

        var search = module.BuildSearchScreen(state, "CUSTOMERS", "Bremen");
        var appended = module.AppendRecord(state, "CUSTOMERS", "Id=C-1004;Company=Retro Works;City=Berlin;Tier=B");
        var updated = module.UpdateRecord(state, "CUSTOMERS", "C-1004", "City", "Leipzig");

        Assert.Contains(search.Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(appended.Content, line => line.Contains("Retro Works", StringComparison.Ordinal));
        Assert.Contains(updated.Content, line => line.Contains("Leipzig", StringComparison.Ordinal));
    }

    [Fact]
    public void DatabaseModule_CanDeleteRecords()
    {
        var module = new DatabaseModule(DatabaseCatalogLoader.LoadDefault());
        var state = module.CreateWorkspaceState();
        module.AppendRecord(state, "CUSTOMERS", "Id=C-1004;Company=Retro Works;City=Berlin;Tier=B");

        var deleted = module.DeleteRecord(state, "CUSTOMERS", "C-1004");

        Assert.DoesNotContain(deleted.Content, line => line.Contains("C-1004", StringComparison.Ordinal));
    }

    [Fact]
    public void DatabaseModule_CanBrowseAcrossCurrentRows()
    {
        var module = new DatabaseModule(DatabaseCatalogLoader.LoadDefault());
        var state = module.CreateWorkspaceState();

        var browse = module.BuildBrowseScreen(state, "CUSTOMERS");
        var next = module.MoveNext(state);
        var previous = module.MovePrevious(state);

        Assert.Contains(browse.Content, line => line.Contains("Current row      : 1 of", StringComparison.Ordinal));
        Assert.Contains(browse.Content, line => line.StartsWith("> ", StringComparison.Ordinal));
        Assert.Contains(next.Content, line => line.Contains("Current row      : 2 of", StringComparison.Ordinal));
        Assert.Contains(previous.Content, line => line.Contains("Current row      : 1 of", StringComparison.Ordinal));
    }
}
