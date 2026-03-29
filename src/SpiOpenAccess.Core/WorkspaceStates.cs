namespace SpiOpenAccess.Core;

public sealed class SpreadsheetWorkspaceState
{
    public decimal Q1 { get; set; } = 182_500m;
    public decimal Q2 { get; set; } = 204_400m;
    public decimal Q3 { get; set; } = 198_225m;
    public decimal Q4 { get; set; } = 227_000m;

    public decimal[] ToArray() => [Q1, Q2, Q3, Q4];

    public void SetQuarter(string quarter, decimal value)
    {
        switch (quarter.Trim().ToUpperInvariant())
        {
            case "Q1":
                Q1 = value;
                break;
            case "Q2":
                Q2 = value;
                break;
            case "Q3":
                Q3 = value;
                break;
            case "Q4":
                Q4 = value;
                break;
            default:
                throw new KeyNotFoundException($"Unknown quarter '{quarter}'.");
        }
    }
}

public sealed class WordProcessorWorkspaceState
{
    public string Title { get; set; } = "OPENACCESS Proposal";
    public string Template { get; set; } = "Executive Letter";
    public List<string> Lines { get; set; } =
    [
        "Dear customer,",
        "please find the current commercial overview attached.",
        "Regards,",
        "ADMIN"
    ];
}

public sealed class MailWorkspaceState
{
    public string To { get; set; } = "SALES";
    public string Subject { get; set; } = "Weekly pipeline update";
    public string Body { get; set; } = "Please send the current pipeline figures before noon.";
}

public sealed class DatabaseWorkspaceState
{
    public List<DatabaseTableState> Tables { get; set; } = [];

    public DatabaseTableState GetTable(string tableName)
    {
        return Tables.FirstOrDefault(table => string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
    }
}

public sealed class DatabaseTableState
{
    public string Name { get; set; } = string.Empty;

    public List<Dictionary<string, string>> Records { get; set; } = [];
}
