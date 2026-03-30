using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Communications;
using SpiOpenAccess.Modules.Database;
using SpiOpenAccess.Modules.Mail;
using SpiOpenAccess.Modules.Programming;
using SpiOpenAccess.Modules.Reporting;
using SpiOpenAccess.Modules.Spreadsheet;
using SpiOpenAccess.Modules.WordProcessing;

namespace SpiOpenAccess.Tests;

public sealed class ModuleCommandScreenTests
{
    private static readonly OfficeWorkspace Workspace = new(
        "OPENACCESS",
        "ADMIN",
        new DateOnly(2026, 3, 29),
        new Dictionary<string, string>());

    [Fact]
    public void Spreadsheet_BuildsOperationalScreens()
    {
        var module = new SpreadsheetModule();
        var state = new SpreadsheetWorkspaceState();

        var recalc = module.BuildRecalcScreen(Workspace, state);
        var grid = module.BuildGridScreen(state);
        var goalSeek = module.BuildGoalSeekScreen("margin 18");
        var printArea = module.BuildPrintAreaScreen("area A1:H56");
        var sum = module.BuildCellSumScreen(state, "A1", "D1");

        Assert.Contains(recalc.Content, line => line.Contains("Workbook consistent", StringComparison.Ordinal));
        Assert.Equal("Worksheet Grid", grid.Title);
        Assert.Contains(goalSeek.Content, line => line.Contains("Suggested value", StringComparison.Ordinal));
        Assert.Contains(printArea.Content, line => line.Contains("Landscape", StringComparison.Ordinal));
        Assert.Contains(sum.Content, line => line.Contains("Range            : A1:D1", StringComparison.Ordinal));
    }

    [Fact]
    public void Word_BuildsEditingScreens()
    {
        var module = new WordProcessingModule();
        var state = new WordProcessorWorkspaceState();

        Assert.Contains(module.BuildNewLetterScreen(Workspace, state).Content, line => line.Contains("lines in draft", StringComparison.Ordinal));
        Assert.Contains(module.BuildMergeScreen().Content, line => line.Contains("Records queued", StringComparison.Ordinal));
        Assert.Contains(module.BuildPreviewScreen("page 1", state).Content, line => line.Contains("Page             : page 1", StringComparison.Ordinal));

        state.MoveCursor(1);
        state.InsertAtCursor("Inserted line");
        state.ReplaceAtCursor("Replaced line");
        state.DeleteAtCursor();
        Assert.True(state.CursorLine >= 1);
    }

    [Fact]
    public void Mail_Comm_Report_Programming_BuildOperationalScreens()
    {
        var mail = new MailModule();
        var communications = new CommunicationsModule();
        var reporting = new ReportingModule();
        var programming = new ProgrammingModule();
        var mailState = new MailWorkspaceState();
        var communicationsState = new CommunicationsWorkspaceState
        {
            CurrentTarget = "HQ",
            IsConnected = true,
            CaptureMode = "on",
            LastTransferFile = "orders.dat"
        };
        var reportingState = new ReportingWorkspaceState
        {
            LastRunReport = "Aging",
            LastRunAt = "2026-03-30 09:30:00",
            ActiveLayout = "Revenue-City",
            OutputHistory = ["Aging-20260330-093000.lst"]
        };
        var programmingState = new ProgrammingWorkspaceState
        {
            ProgramName = "sample.pro",
            LastRunAt = "2026-03-30 09:45:00"
        };
        var databaseState = new DatabaseModule(DatabaseCatalogLoader.LoadDefault()).CreateWorkspaceState();

        Assert.Contains(mail.BuildComposeScreen(Workspace, mailState).Content, line => line.Contains("Weekly pipeline update", StringComparison.Ordinal));
        mailState.SentItems.Add(new MailSentItemState { To = "FINANCE", Subject = "Cash status", Body = "Need current cash figures.", SentAt = "2026-03-29 10:15:00" });
        Assert.Contains(mail.BuildSentItemsScreen(mailState).Content, line => line.Contains("Cash status", StringComparison.Ordinal));
        Assert.Contains(mail.BuildRoutingRulesScreen().Content, line => line.Contains("Rule 1", StringComparison.Ordinal));
        Assert.Contains(mail.BuildMessageScreen("OPS-142").Content, line => line.Contains("Quarter close checklist", StringComparison.Ordinal));

        Assert.Contains(communications.BuildHomeScreen(Workspace, communicationsState).Content, line => line.Contains("Current target", StringComparison.Ordinal));
        Assert.Contains(communications.BuildDialScreen(communicationsState, "HQ").Content, line => line.Contains("Carrier detected", StringComparison.Ordinal));
        Assert.Contains(communications.BuildSendScreen(communicationsState, "orders.dat").Content, line => line.Contains("Transfer complete", StringComparison.Ordinal));
        Assert.Contains(communications.BuildCaptureScreen(communicationsState, "on").Content, line => line.Contains("Capture armed", StringComparison.Ordinal));

        Assert.Contains(reporting.BuildHomeScreen(Workspace, reportingState).Content, line => line.Contains("Active layout", StringComparison.Ordinal));
        Assert.Contains(reporting.BuildRunScreen(reportingState, "Aging", databaseState).Content, line => line.Contains("Completed", StringComparison.Ordinal));
        Assert.Contains(reporting.BuildScheduleScreen(reportingState, "Evening Finance").Content, line => line.Contains("Weekdays", StringComparison.Ordinal));
        Assert.Contains(reporting.BuildDesignScreen(reportingState, "Revenue-City").Content, line => line.Contains("Sections", StringComparison.Ordinal));

        Assert.Contains(programming.BuildHomeScreen(Workspace, programmingState).Content, line => line.Contains("Program", StringComparison.Ordinal));
        Assert.Contains(programming.BuildRunScreen(programmingState, "sample.pro").Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(programming.BuildVariablesScreen(programmingState).Content, line => line.Contains("openInvoices", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(programming.BuildCompileScreen(programmingState, "nightly.pro").Content, line => line.Contains("Compile successful", StringComparison.Ordinal));
        Assert.Contains(programming.BuildSourceScreen(programmingState).Content, line => line.Contains("LET company", StringComparison.Ordinal));
    }
}
