using SpiOpenAccess.Core;

namespace SpiOpenAccess.App;

public sealed class AppSessionState
{
    public DatabaseWorkspaceState Database { get; set; } = new();

    public SpreadsheetWorkspaceState Spreadsheet { get; set; } = new();

    public WordProcessorWorkspaceState Word { get; set; } = new();

    public MailWorkspaceState Mail { get; set; } = new();

    public CommunicationsWorkspaceState Communications { get; set; } = new();

    public ReportingWorkspaceState Reporting { get; set; } = new();

    public ProgrammingWorkspaceState Programming { get; set; } = new();
}
