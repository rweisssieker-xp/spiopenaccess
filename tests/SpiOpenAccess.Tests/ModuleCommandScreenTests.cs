using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Communications;
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
        var goalSeek = module.BuildGoalSeekScreen("margin 18");
        var printArea = module.BuildPrintAreaScreen("area A1:H56");

        Assert.Contains(recalc.Content, line => line.Contains("Workbook consistent", StringComparison.Ordinal));
        Assert.Contains(goalSeek.Content, line => line.Contains("Suggested value", StringComparison.Ordinal));
        Assert.Contains(printArea.Content, line => line.Contains("Landscape", StringComparison.Ordinal));
    }

    [Fact]
    public void Word_BuildsEditingScreens()
    {
        var module = new WordProcessingModule();
        var state = new WordProcessorWorkspaceState();

        Assert.Contains(module.BuildNewLetterScreen(Workspace, state).Content, line => line.Contains("lines in draft", StringComparison.Ordinal));
        Assert.Contains(module.BuildMergeScreen().Content, line => line.Contains("Records queued", StringComparison.Ordinal));
        Assert.Contains(module.BuildPreviewScreen("page 1", state).Content, line => line.Contains("Page             : page 1", StringComparison.Ordinal));
    }

    [Fact]
    public void Mail_Comm_Report_Programming_BuildOperationalScreens()
    {
        var mail = new MailModule();
        var communications = new CommunicationsModule();
        var reporting = new ReportingModule();
        var programming = new ProgrammingModule();
        var mailState = new MailWorkspaceState();

        Assert.Contains(mail.BuildComposeScreen(Workspace, mailState).Content, line => line.Contains("Weekly pipeline update", StringComparison.Ordinal));
        Assert.Contains(mail.BuildRoutingRulesScreen().Content, line => line.Contains("Rule 1", StringComparison.Ordinal));
        Assert.Contains(mail.BuildMessageScreen("OPS-142").Content, line => line.Contains("Quarter close checklist", StringComparison.Ordinal));

        Assert.Contains(communications.BuildDialScreen("HQ").Content, line => line.Contains("Carrier detected", StringComparison.Ordinal));
        Assert.Contains(communications.BuildSendScreen("orders.dat").Content, line => line.Contains("Transfer complete", StringComparison.Ordinal));
        Assert.Contains(communications.BuildCaptureScreen("on").Content, line => line.Contains("Capture armed", StringComparison.Ordinal));

        Assert.Contains(reporting.BuildRunScreen("Aging").Content, line => line.Contains("Completed", StringComparison.Ordinal));
        Assert.Contains(reporting.BuildScheduleScreen("Evening Finance").Content, line => line.Contains("Weekdays", StringComparison.Ordinal));
        Assert.Contains(reporting.BuildDesignScreen("Revenue-City").Content, line => line.Contains("Sections", StringComparison.Ordinal));

        Assert.Contains(programming.BuildRunScreen("sample.pro").Content, line => line.Contains("Nordwest Handel", StringComparison.Ordinal));
        Assert.Contains(programming.BuildVariablesScreen().Content, line => line.Contains("openInvoices", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(programming.BuildCompileScreen("nightly.pro").Content, line => line.Contains("Compile successful", StringComparison.Ordinal));
    }
}
