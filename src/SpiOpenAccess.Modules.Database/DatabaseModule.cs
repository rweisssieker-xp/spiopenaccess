using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Database;

public sealed class DatabaseModule : IOfficeModule
{
    private readonly DatabaseCatalog _catalog;

    public DatabaseModule()
        : this(DatabaseCatalogLoader.LoadDefault())
    {
    }

    public DatabaseModule(DatabaseCatalog catalog)
    {
        _catalog = catalog;
    }

    public ModuleInfo Info { get; } = new(
        "db",
        "Desktop Database",
        "Relational tables, forms, indexes, and shared multi-user workflows.",
        "Data",
        ["Tables", "Forms", "Reports", "Locking", "Import/Export"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        return BuildHomeScreen(workspace, CreateWorkspaceState());
    }

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace, DatabaseWorkspaceState state)
    {
        var customersTable = _catalog.Tables.FirstOrDefault(table =>
            string.Equals(table.Name, "CUSTOMERS", StringComparison.OrdinalIgnoreCase));
        var firstCustomer = customersTable is null ? null : GetRecords(state, customersTable.Name).FirstOrDefault();
        var content = new List<string>
        {
            $"Workspace        : {workspace.Name}",
            $"Snapshot         : {workspace.SnapshotDate:yyyy-MM-dd}",
            "Mode             : Shared network workspace",
            "Record locking   : Optimistic with explicit save",
            $"Tables           : {_catalog.Tables.Count}",
            $"Forms            : {_catalog.Forms.Count}",
            $"Reports          : {_catalog.Reports.Count}",
            "Catalog          :"
        };

        content.AddRange(_catalog.Tables.Select(table =>
            $"  {table.Name,-12} {GetRecords(state, table.Name).Count,4} rows  key={table.KeyField,-10} cols=[{string.Join(", ", table.Columns.Select(column => column.Name))}]"));
        content.Add("Form layouts      :");
        content.AddRange(_catalog.Forms.Select(form =>
            $"  {form.Name,-14} table={form.Table,-10} fields=[{string.Join(", ", form.Fields)}]"));
        content.Add("Report layouts    :");
        content.AddRange(_catalog.Reports.Select(report =>
            $"  {report.Name,-14} table={report.Table,-10} group={report.GroupBy,-10} sums=[{string.Join(", ", report.Summaries)}]"));

        if (firstCustomer is not null)
        {
            content.Add("Current record    :");
            content.AddRange(firstCustomer.Select(field => $"  {field.Key,-16} {field.Value}"));
        }

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["open CUSTOMERS", "edit CUSTOMER_CARD", "run AGING"]);
    }

    public DatabaseWorkspaceState CreateWorkspaceState()
    {
        return new DatabaseWorkspaceState
        {
            Tables = _catalog.Tables.Select(table => new DatabaseTableState
            {
                Name = table.Name,
                Records = table.Records
                    .Select(record => new Dictionary<string, string>(record, StringComparer.OrdinalIgnoreCase))
                    .ToList()
            }).ToList()
        };
    }

    public DatabaseTable? FindTable(string tableName)
    {
        return _catalog.Tables.FirstOrDefault(table =>
            string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase));
    }

    public IReadOnlyList<FormDefinition> GetFormsForTable(string tableName)
    {
        return _catalog.Forms
            .Where(form => string.Equals(form.Table, tableName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
    }

    public IReadOnlyList<ReportDefinition> GetReportsForTable(string tableName)
    {
        return _catalog.Reports
            .Where(report => string.Equals(report.Table, tableName, StringComparison.OrdinalIgnoreCase))
            .ToArray();
    }

    public ModuleScreen BuildTableScreen(string tableName)
    {
        return BuildTableScreen(CreateWorkspaceState(), tableName);
    }

    public ModuleScreen BuildTableScreen(DatabaseWorkspaceState state, string tableName)
    {
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var records = GetRecords(state, table.Name);
        var firstRecord = records.FirstOrDefault();
        var content = new List<string>
        {
            $"Table            : {table.Name}",
            $"Key field        : {table.KeyField}",
            $"Columns          : {string.Join(", ", table.Columns.Select(column => $"{column.Name}:{column.Type}({column.Width})"))}",
            $"Record count     : {records.Count}",
            "Rows             :"
        };

        foreach (var record in records.Take(10))
        {
            content.Add($"  {string.Join(" | ", record.Select(field => $"{field.Key}={field.Value}"))}");
        }

        if (firstRecord is not null)
        {
            content.Add("Current pointer   :");
            content.AddRange(firstRecord.Select(field => $"  {field.Key,-16} {field.Value}"));
        }

        return ModuleScreen.Create(
            $"Table {table.Name}",
            "Table view with live record preview.",
            content,
            [$"find {table.Name} <term>", $"append {table.Name} field=value;...", $"update {table.Name} key field value", $"delete {table.Name} key"]);
    }

    public ModuleScreen BuildFormScreen(string formName)
    {
        return BuildFormScreen(CreateWorkspaceState(), formName);
    }

    public ModuleScreen BuildFormScreen(DatabaseWorkspaceState state, string formName)
    {
        var form = _catalog.Forms.FirstOrDefault(candidate =>
            string.Equals(candidate.Name, formName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown form '{formName}'.");
        var table = FindTable(form.Table) ?? throw new KeyNotFoundException($"Unknown table '{form.Table}'.");
        var sampleRecord = GetRecords(state, table.Name).FirstOrDefault();

        var content = new List<string>
        {
            $"Form             : {form.Name}",
            $"Bound table      : {form.Table}",
            $"Field count      : {form.Fields.Count}",
            "Layout fields    :"
        };
        content.AddRange(form.Fields.Select(field => $"  {field}"));

        if (sampleRecord is not null)
        {
            content.Add("Sample values     :");
            content.AddRange(form.Fields.Select(field =>
                $"  {field,-16} {sampleRecord.GetValueOrDefault(field, "<empty>")}"));
        }

        return ModuleScreen.Create(
            $"Form {form.Name}",
            "Form definition with bound fields.",
            content,
            [$"open {form.Table}", $"run {GetPrimaryReportForTable(form.Table)?.Name ?? "REPORT"}", "back"]);
    }

    public ModuleScreen BuildReportScreen(string reportName)
    {
        return BuildReportScreen(CreateWorkspaceState(), reportName);
    }

    public ModuleScreen BuildReportScreen(DatabaseWorkspaceState state, string reportName)
    {
        var report = _catalog.Reports.FirstOrDefault(candidate =>
            string.Equals(candidate.Name, reportName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown report '{reportName}'.");
        var table = FindTable(report.Table) ?? throw new KeyNotFoundException($"Unknown table '{report.Table}'.");

        var records = GetRecords(state, table.Name);
        var content = new List<string>
        {
            $"Report           : {report.Name}",
            $"Source table     : {report.Table}",
            $"Group by         : {report.GroupBy}",
            $"Summaries        : {string.Join(", ", report.Summaries)}",
            $"Input rows       : {records.Count}",
            "Preview          :"
        };

        foreach (var grouping in records.GroupBy(record => record.GetValueOrDefault(report.GroupBy, "<null>")))
        {
            content.Add($"  {grouping.Key,-16} count={grouping.Count()}");
        }

        return ModuleScreen.Create(
            $"Report {report.Name}",
            "Report definition with grouped preview output.",
            content,
            [$"open {report.Table}", $"edit {GetPrimaryFormForTable(report.Table)?.Name ?? "FORM"}", "back"]);
    }

    public FormDefinition? GetPrimaryFormForTable(string tableName)
    {
        return GetFormsForTable(tableName).FirstOrDefault();
    }

    public ReportDefinition? GetPrimaryReportForTable(string tableName)
    {
        return GetReportsForTable(tableName).FirstOrDefault();
    }

    public ModuleScreen BuildSearchScreen(DatabaseWorkspaceState state, string tableName, string searchTerm)
    {
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var matches = GetRecords(state, table.Name)
            .Where(record => record.Values.Any(value => value.Contains(searchTerm, StringComparison.OrdinalIgnoreCase)))
            .ToList();

        var content = new List<string>
        {
            $"Table            : {table.Name}",
            $"Search term      : {searchTerm}",
            $"Matches          : {matches.Count}",
            "Results          :"
        };
        content.AddRange(matches.Select(record => $"  {string.Join(" | ", record.Select(field => $"{field.Key}={field.Value}"))}"));
        if (matches.Count == 0)
        {
            content.Add("  <none>");
        }

        return ModuleScreen.Create(
            $"Find {table.Name}",
            "Search across visible table fields.",
            content,
            [$"open {table.Name}", $"append {table.Name} field=value;...", "back"]);
    }

    public ModuleScreen AppendRecord(DatabaseWorkspaceState state, string tableName, string assignments)
    {
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var record = ParseAssignments(assignments, table.Columns.Select(column => column.Name));
        var keyField = table.KeyField;
        if (!record.ContainsKey(keyField))
        {
            throw new InvalidOperationException($"Missing key field '{keyField}'.");
        }

        GetRecords(state, table.Name).Add(new Dictionary<string, string>(record, StringComparer.OrdinalIgnoreCase));
        return BuildTableScreen(state, table.Name);
    }

    public ModuleScreen UpdateRecord(DatabaseWorkspaceState state, string tableName, string keyValue, string fieldName, string newValue)
    {
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var keyField = table.KeyField;
        var record = GetRecords(state, table.Name)
            .FirstOrDefault(candidate => string.Equals(candidate.GetValueOrDefault(keyField), keyValue, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown record '{keyValue}' in table '{tableName}'.");

        if (!table.Columns.Any(column => string.Equals(column.Name, fieldName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new KeyNotFoundException($"Unknown field '{fieldName}' in table '{tableName}'.");
        }

        record[fieldName] = newValue;
        return BuildTableScreen(state, table.Name);
    }

    public ModuleScreen DeleteRecord(DatabaseWorkspaceState state, string tableName, string keyValue)
    {
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var keyField = table.KeyField;
        var records = GetRecords(state, table.Name);
        var record = records
            .FirstOrDefault(candidate => string.Equals(candidate.GetValueOrDefault(keyField), keyValue, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown record '{keyValue}' in table '{tableName}'.");

        records.Remove(record);
        return BuildTableScreen(state, table.Name);
    }

    private static List<Dictionary<string, string>> GetRecords(DatabaseWorkspaceState state, string tableName)
    {
        return state.GetTable(tableName).Records;
    }

    private static Dictionary<string, string> ParseAssignments(string assignments, IEnumerable<string> validColumns)
    {
        var valid = new HashSet<string>(validColumns, StringComparer.OrdinalIgnoreCase);
        var record = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var pairs = assignments.Split(';', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        foreach (var pair in pairs)
        {
            var parts = pair.Split('=', 2, StringSplitOptions.TrimEntries);
            if (parts.Length != 2)
            {
                throw new InvalidOperationException("Assignments must use field=value;field=value.");
            }

            if (!valid.Contains(parts[0]))
            {
                throw new KeyNotFoundException($"Unknown field '{parts[0]}'.");
            }

            record[parts[0]] = parts[1];
        }

        return record;
    }
}
