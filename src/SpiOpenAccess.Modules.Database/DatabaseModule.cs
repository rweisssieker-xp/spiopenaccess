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
        "Relationale Tabellen, Masken, Indizes und Mehrbenutzerbetrieb.",
        "Data",
        ["Tables", "Forms", "Reports", "Locking", "Import/Export"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var customersTable = _catalog.Tables.FirstOrDefault(table =>
            string.Equals(table.Name, "CUSTOMERS", StringComparison.OrdinalIgnoreCase));

        var firstCustomer = customersTable?.Records.FirstOrDefault();
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
            $"  {table.Name,-12} {table.Records.Count,4} rows  key={table.KeyField,-10} cols=[{string.Join(", ", table.Columns.Select(column => column.Name))}]"));
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
        var table = FindTable(tableName) ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
        var content = new List<string>
        {
            $"Table            : {table.Name}",
            $"Key field        : {table.KeyField}",
            $"Columns          : {string.Join(", ", table.Columns.Select(column => $"{column.Name}:{column.Type}({column.Width})"))}",
            $"Record count     : {table.Records.Count}",
            "Rows             :"
        };

        foreach (var record in table.Records.Take(10))
        {
            content.Add($"  {string.Join(" | ", record.Select(field => $"{field.Key}={field.Value}"))}");
        }

        return ModuleScreen.Create(
            $"Table {table.Name}",
            "Tabellenansicht mit Datensatzvorschau.",
            content,
            [$"edit {GetPrimaryFormForTable(table.Name)?.Name ?? "FORM"}", $"run {GetPrimaryReportForTable(table.Name)?.Name ?? "REPORT"}", "back"]);
    }

    public ModuleScreen BuildFormScreen(string formName)
    {
        var form = _catalog.Forms.FirstOrDefault(candidate =>
            string.Equals(candidate.Name, formName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown form '{formName}'.");
        var table = FindTable(form.Table) ?? throw new KeyNotFoundException($"Unknown table '{form.Table}'.");
        var sampleRecord = table.Records.FirstOrDefault();

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
            "Maskendefinition mit gebundenen Feldern.",
            content,
            [$"open {form.Table}", $"run {GetPrimaryReportForTable(form.Table)?.Name ?? "REPORT"}", "back"]);
    }

    public ModuleScreen BuildReportScreen(string reportName)
    {
        var report = _catalog.Reports.FirstOrDefault(candidate =>
            string.Equals(candidate.Name, reportName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown report '{reportName}'.");
        var table = FindTable(report.Table) ?? throw new KeyNotFoundException($"Unknown table '{report.Table}'.");

        var content = new List<string>
        {
            $"Report           : {report.Name}",
            $"Source table     : {report.Table}",
            $"Group by         : {report.GroupBy}",
            $"Summaries        : {string.Join(", ", report.Summaries)}",
            $"Input rows       : {table.Records.Count}",
            "Preview          :"
        };

        foreach (var grouping in table.Records.GroupBy(record => record.GetValueOrDefault(report.GroupBy, "<null>")))
        {
            content.Add($"  {grouping.Key,-16} count={grouping.Count()}");
        }

        return ModuleScreen.Create(
            $"Report {report.Name}",
            "Reportdefinition mit einfacher Gruppen-Vorschau.",
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
}
