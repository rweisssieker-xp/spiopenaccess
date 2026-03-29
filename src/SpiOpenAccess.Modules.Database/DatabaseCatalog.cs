using System.Text.Json.Serialization;

namespace SpiOpenAccess.Modules.Database;

public sealed record DatabaseCatalog(
    [property: JsonPropertyName("tables")] IReadOnlyList<DatabaseTable> Tables,
    [property: JsonPropertyName("forms")] IReadOnlyList<FormDefinition> Forms,
    [property: JsonPropertyName("reports")] IReadOnlyList<ReportDefinition> Reports);

public sealed record DatabaseTable(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("keyField")] string KeyField,
    [property: JsonPropertyName("columns")] IReadOnlyList<DatabaseColumn> Columns,
    [property: JsonPropertyName("records")] IReadOnlyList<IReadOnlyDictionary<string, string>> Records);

public sealed record DatabaseColumn(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("type")] string Type,
    [property: JsonPropertyName("width")] int Width);

public sealed record FormDefinition(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("table")] string Table,
    [property: JsonPropertyName("fields")] IReadOnlyList<string> Fields);

public sealed record ReportDefinition(
    [property: JsonPropertyName("name")] string Name,
    [property: JsonPropertyName("table")] string Table,
    [property: JsonPropertyName("groupBy")] string GroupBy,
    [property: JsonPropertyName("summaries")] IReadOnlyList<string> Summaries);
