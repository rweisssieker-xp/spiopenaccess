using SpiOpenAccess.Core;
using SpiOpenAccess.Infrastructure;
using SpiOpenAccess.Modules.Communications;
using SpiOpenAccess.Modules.Database;
using SpiOpenAccess.Modules.Mail;
using SpiOpenAccess.Modules.Programming;
using SpiOpenAccess.Modules.Reporting;
using SpiOpenAccess.Modules.Spreadsheet;
using SpiOpenAccess.Modules.WordProcessing;
using SpiOpenAccess.App;

const int FrameWidth = 80;
const int VisibleBodyLines = 17;
var suite = OfficeSuiteFactory.CreateDefault();
var sessionStore = new AppSessionStore(Directory.GetCurrentDirectory());
var sessionState = sessionStore.Load();
SeedDatabaseStateIfNeeded(suite, sessionState);

if (args.Length > 0)
{
    var selector = string.Join(' ', args);
    RenderSelectedModule(suite, selector, sessionState);
    return;
}

RenderShell(suite, sessionState, sessionStore);

static void RenderShell(OfficeSuite suite, AppSessionState sessionState, AppSessionStore sessionStore)
{
    var activeModule = suite.FindModule("db") ?? suite.Modules[0];
    var currentScreen = BuildHomeScreen(activeModule, suite.Workspace, sessionState);
    var status = "F1 Menu  F2 Modules  F3 Home  F10 Exit";

    while (true)
    {
        RenderWorkspace(suite, activeModule, currentScreen, status);
        var input = ReadCommand(activeModule);
        if (string.IsNullOrWhiteSpace(input))
        {
            status = "No command entered.";
            continue;
        }

        input = NormalizeCommandAlias(input);

        if (string.Equals(input, "__NAV_LEFT__", StringComparison.Ordinal))
        {
            activeModule = CycleModule(suite, activeModule, -1);
            currentScreen = BuildHomeScreen(activeModule, suite.Workspace, sessionState);
            status = $"Switched to {activeModule.Info.DisplayName}.";
            continue;
        }

        if (string.Equals(input, "__NAV_RIGHT__", StringComparison.Ordinal))
        {
            activeModule = CycleModule(suite, activeModule, 1);
            currentScreen = BuildHomeScreen(activeModule, suite.Workspace, sessionState);
            status = $"Switched to {activeModule.Info.DisplayName}.";
            continue;
        }

        if (IsQuit(input))
        {
            return;
        }

        if (string.Equals(input, "menu", StringComparison.OrdinalIgnoreCase))
        {
            var selectedModule = RenderModuleMenuAndReadSelection(suite, activeModule);
            if (selectedModule is not null)
            {
                activeModule = selectedModule;
                currentScreen = BuildHomeScreen(activeModule, suite.Workspace, sessionState);
                status = $"Switched to {activeModule.Info.DisplayName}.";
            }
            else
            {
                status = "Module switch cancelled.";
            }

            continue;
        }

        if (TryResolveModuleSelection(suite, input, out var switchedModule))
        {
            activeModule = switchedModule!;
            currentScreen = BuildHomeScreen(activeModule, suite.Workspace, sessionState);
            status = $"Switched to {activeModule.Info.DisplayName}.";
            continue;
        }

        if (TryBuildCommandScreen(activeModule, suite.Workspace, sessionState, input, out var commandScreen, out var error, out var persistState))
        {
            currentScreen = commandScreen!;
            status = $"Executed: {input}";
            if (persistState)
            {
                sessionStore.Save(sessionState);
                status += " (saved)";
            }
            continue;
        }

        status = error ?? $"Unknown command: {input}";
    }
}

static void RenderSelectedModule(OfficeSuite suite, string selector, AppSessionState sessionState)
{
    var module = suite.FindModule(selector);
    if (module is null)
    {
        Console.Error.WriteLine($"Unknown module '{selector}'.");
        Environment.ExitCode = 1;
        return;
    }

    RenderWorkspace(
        suite,
        module,
        BuildHomeScreen(module, suite.Workspace, sessionState),
        "Direct module view.");
}

static ModuleScreen BuildHomeScreen(IOfficeModule module, OfficeWorkspace workspace, AppSessionState sessionState)
{
    return module switch
    {
        DatabaseModule database => database.BuildHomeScreen(workspace, sessionState.Database),
        SpreadsheetModule spreadsheet => spreadsheet.BuildHomeScreen(workspace, sessionState.Spreadsheet),
        WordProcessingModule word => word.BuildHomeScreen(workspace, sessionState.Word),
        MailModule mail => mail.BuildHomeScreen(workspace, sessionState.Mail),
        _ => module.BuildHomeScreen(workspace)
    };
}

static void SeedDatabaseStateIfNeeded(OfficeSuite suite, AppSessionState sessionState)
{
    if (sessionState.Database.Tables.Count > 0)
    {
        return;
    }

    if (suite.FindModule("db") is DatabaseModule databaseModule)
    {
        sessionState.Database = databaseModule.CreateWorkspaceState();
    }
}

static void RenderWorkspace(OfficeSuite suite, IOfficeModule module, ModuleScreen screen, string status)
{
    SafeClear();
    RenderTitleBar(suite, module);
    RenderMenuBar(suite, module);
    RenderScreen(screen);
    RenderCommandHints();
    RenderStatusBar(status);
}

static void RenderTitleBar(OfficeSuite suite, IOfficeModule module)
{
    Console.ForegroundColor = ConsoleColor.Black;
    Console.BackgroundColor = ConsoleColor.Gray;
    Console.WriteLine(Fit($" SPI OPEN ACCESS REBUILD  VER {suite.Version,-6} WS:{suite.Workspace.Name,-10} USER:{suite.Workspace.Owner,-8} MOD:{module.Info.DisplayName} ", FrameWidth));
    Console.ResetColor();
}

static void RenderMenuBar(OfficeSuite suite, IOfficeModule activeModule)
{
    var items = suite.Modules.Select(module =>
    {
        var label = GetMenuLabel(module.Info.Id);
        return module.Info.Id == activeModule.Info.Id ? $"[{label}]" : $" {label} ";
    });

    Console.ForegroundColor = ConsoleColor.White;
    Console.BackgroundColor = ConsoleColor.DarkBlue;
    Console.WriteLine(Fit(string.Join(" ", items), FrameWidth));
    Console.ResetColor();
}

static void RenderScreen(ModuleScreen screen)
{
    var innerWidth = FrameWidth - 4;
    var lines = new List<string>
    {
        screen.Summary,
        string.Empty
    };

    lines.AddRange(screen.Content);

    if (screen.Commands.Count > 0)
    {
        lines.Add(string.Empty);
        lines.Add("Commands:");
        lines.AddRange(screen.Commands.Select(command => $"  {command}"));
    }

    var title = $"[ {screen.Title.ToUpperInvariant()} ]";
    Console.WriteLine("+" + title + new string('-', Math.Max(0, FrameWidth - 3 - title.Length)) + "+");
    foreach (var line in lines.SelectMany(line => Wrap(line, innerWidth)))
    {
        Console.WriteLine($"| {Fit(line, innerWidth)} |");
    }

    var usedLines = lines.SelectMany(line => Wrap(line, innerWidth)).Count();
    for (var index = usedLines; index < VisibleBodyLines; index++)
    {
        Console.WriteLine($"| {new string(' ', innerWidth)} |");
    }

    Console.WriteLine("+" + new string('-', FrameWidth - 2) + "+");
}

static void RenderCommandHints()
{
    Console.ForegroundColor = ConsoleColor.Black;
    Console.BackgroundColor = ConsoleColor.Gray;
    Console.WriteLine(Fit(" Command: menu | use <module> | open <table> | edit <form> | run <report> | back ", FrameWidth));
    Console.ResetColor();
}

static void RenderStatusBar(string status)
{
    Console.ForegroundColor = ConsoleColor.Yellow;
    Console.BackgroundColor = ConsoleColor.DarkBlue;
    Console.WriteLine(Fit(" F1=Menu  F2=Modules  F3=Home  F5=Open  F6=Edit  F7=Run  <- -> Tabs  F10=Exit ", FrameWidth));
    Console.ForegroundColor = ConsoleColor.Black;
    Console.BackgroundColor = ConsoleColor.Gray;
    Console.WriteLine(Fit($" Status: {status} ", FrameWidth));
    Console.ResetColor();
}

static IOfficeModule? RenderModuleMenuAndReadSelection(OfficeSuite suite, IOfficeModule activeModule)
{
    SafeClear();
    RenderTitleBar(suite, activeModule);
    Console.ForegroundColor = ConsoleColor.White;
    Console.BackgroundColor = ConsoleColor.DarkBlue;
    Console.WriteLine(Fit(" MODULE SELECTION ", FrameWidth));
    Console.ResetColor();
    Console.WriteLine("+" + new string('=', FrameWidth - 2) + "+");

    for (var index = 0; index < suite.Modules.Count; index++)
    {
        var module = suite.Modules[index];
        var marker = module.Info.Id == activeModule.Info.Id ? "*" : " ";
        var line = $"{marker} {index + 1,2}. {module.Info.DisplayName,-18} ({module.Info.Id}) {module.Info.Summary}";
        Console.WriteLine($"| {Fit(line, FrameWidth - 4)} |");
    }

    for (var index = suite.Modules.Count; index < 12; index++)
    {
        Console.WriteLine($"| {new string(' ', FrameWidth - 4)} |");
    }

    Console.WriteLine("+" + new string('=', FrameWidth - 2) + "+");
    RenderStatusBar("Enter number, id or name. Empty input cancels.");
    Console.Write("Select> ");
    var input = Console.ReadLine()?.Trim();
    if (string.IsNullOrWhiteSpace(input))
    {
        return null;
    }

    return TryResolveModuleSelection(suite, input, out var moduleSelection) ? moduleSelection : null;
}

static bool TryResolveModuleSelection(OfficeSuite suite, string input, out IOfficeModule? module)
{
    module = null;

    if (input.StartsWith("use ", StringComparison.OrdinalIgnoreCase))
    {
        input = input[4..].Trim();
    }

    if (int.TryParse(input, out var selectedIndex) && selectedIndex >= 1 && selectedIndex <= suite.Modules.Count)
    {
        module = suite.Modules[selectedIndex - 1];
        return true;
    }

    module = suite.FindModule(input);
    return module is not null;
}

static bool TryBuildCommandScreen(
    IOfficeModule activeModule,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string input,
    out ModuleScreen? screen,
    out string? error,
    out bool persistState)
{
    screen = null;
    error = null;
    persistState = false;

    if (string.Equals(input, "back", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(input, "home", StringComparison.OrdinalIgnoreCase))
    {
        screen = BuildHomeScreen(activeModule, workspace, sessionState);
        return true;
    }

    var commandParts = input.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (commandParts.Length == 0)
    {
        return false;
    }

    if (commandParts.Length == 1)
    {
        try
        {
            screen = activeModule switch
            {
                SpreadsheetModule spreadsheet when commandParts[0].Equals("recalc", StringComparison.OrdinalIgnoreCase)
                    => spreadsheet.BuildRecalcScreen(workspace, sessionState.Spreadsheet),
                MailModule mail when commandParts[0].Equals("compose", StringComparison.OrdinalIgnoreCase)
                    => mail.BuildComposeScreen(workspace, sessionState.Mail),
                MailModule mail when commandParts[0].Equals("send", StringComparison.OrdinalIgnoreCase)
                    => HandleMailSend(mail, sessionState, out persistState),
                MailModule mail when commandParts[0].Equals("sent", StringComparison.OrdinalIgnoreCase)
                    => mail.BuildSentItemsScreen(sessionState.Mail),
                MailModule mail when string.Equals(commandParts[0], "route", StringComparison.OrdinalIgnoreCase)
                    => mail.BuildRoutingRulesScreen(),
                CommunicationsModule communications when commandParts[0].Equals("capture", StringComparison.OrdinalIgnoreCase)
                    => communications.BuildCaptureScreen("on"),
                ReportingModule reporting when commandParts[0].Equals("run", StringComparison.OrdinalIgnoreCase)
                    => reporting.BuildRunScreen("Aging", sessionState.Database),
                ProgrammingModule programming when commandParts[0].Equals("list", StringComparison.OrdinalIgnoreCase)
                    => programming.BuildVariablesScreen(),
                WordProcessingModule word when commandParts[0].Equals("delete-line", StringComparison.OrdinalIgnoreCase)
                    => HandleWordDeleteLine(word, workspace, sessionState, out persistState),
                _ => null
            };
        }
        catch (Exception exception) when (exception is KeyNotFoundException or InvalidOperationException)
        {
            error = exception.Message;
            return false;
        }

        return screen is not null;
    }

    try
    {
        screen = activeModule switch
        {
            DatabaseModule databaseModule => commandParts[0].ToLowerInvariant() switch
            {
                "open" => databaseModule.BuildTableScreen(sessionState.Database, commandParts[1]),
                "edit" => databaseModule.BuildFormScreen(sessionState.Database, commandParts[1]),
                "run" => databaseModule.BuildReportScreen(sessionState.Database, commandParts[1]),
                "find" => HandleDatabaseFind(databaseModule, sessionState, commandParts[1]),
                "append" => HandleDatabaseAppend(databaseModule, sessionState, commandParts[1], out persistState),
                "update" => HandleDatabaseUpdate(databaseModule, sessionState, commandParts[1], out persistState),
                "delete" => HandleDatabaseDelete(databaseModule, sessionState, commandParts[1], out persistState),
                _ => null
            },
            SpreadsheetModule spreadsheet => commandParts[0].ToLowerInvariant() switch
            {
                "goal-seek" => spreadsheet.BuildGoalSeekScreen(commandParts[1]),
                "print" => spreadsheet.BuildPrintAreaScreen(commandParts[1]),
                "set" => HandleSpreadsheetSet(spreadsheet, workspace, sessionState, commandParts[1], out persistState),
                _ => null
            },
            WordProcessingModule word => commandParts[0].ToLowerInvariant() switch
            {
                "new" => HandleWordNew(word, workspace, sessionState, out persistState),
                "merge" => word.BuildMergeScreen(),
                "preview" => word.BuildPreviewScreen(commandParts[1], sessionState.Word),
                "type" => HandleWordType(word, workspace, sessionState, commandParts[1], out persistState),
                "title" => HandleWordTitle(word, workspace, sessionState, commandParts[1], out persistState),
                "delete-line" => HandleWordDeleteLine(word, workspace, sessionState, out persistState),
                _ => null
            },
            MailModule mail => commandParts[0].ToLowerInvariant() switch
            {
                "open" => mail.BuildMessageScreen(commandParts[1]),
                "route" => mail.BuildRoutingRulesScreen(),
                "to" => HandleMailTo(mail, workspace, sessionState, commandParts[1], out persistState),
                "subject" => HandleMailSubject(mail, workspace, sessionState, commandParts[1], out persistState),
                "body" => HandleMailBody(mail, workspace, sessionState, commandParts[1], out persistState),
                "send" => HandleMailSend(mail, sessionState, out persistState),
                _ => null
            },
            CommunicationsModule communications => commandParts[0].ToLowerInvariant() switch
            {
                "dial" => communications.BuildDialScreen(commandParts[1]),
                "send" => communications.BuildSendScreen(commandParts[1]),
                "capture" => communications.BuildCaptureScreen(commandParts[1]),
                _ => null
            },
            ReportingModule reporting => commandParts[0].ToLowerInvariant() switch
            {
                "run" => reporting.BuildRunScreen(commandParts[1], sessionState.Database),
                "schedule" => reporting.BuildScheduleScreen(commandParts[1]),
                "design" => reporting.BuildDesignScreen(commandParts[1]),
                _ => null
            },
            ProgrammingModule programming => commandParts[0].ToLowerInvariant() switch
            {
                "run" => programming.BuildRunScreen(commandParts[1]),
                "list" => programming.BuildVariablesScreen(),
                "compile" => programming.BuildCompileScreen(commandParts[1]),
                _ => null
            },
            _ => null
        };
    }
    catch (Exception exception) when (exception is KeyNotFoundException or InvalidOperationException)
    {
        error = exception.Message;
        return false;
    }

    return screen is not null;
}

static ModuleScreen HandleSpreadsheetSet(
    SpreadsheetModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string argument,
    out bool persistState)
{
    var parts = argument.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2 || !decimal.TryParse(parts[1], out var value))
    {
        throw new InvalidOperationException("Usage: set Q1 190000");
    }

    sessionState.Spreadsheet.SetQuarter(parts[0], value);
    persistState = true;
    return module.BuildHomeScreen(workspace, sessionState.Spreadsheet);
}

static ModuleScreen HandleDatabaseFind(
    DatabaseModule module,
    AppSessionState sessionState,
    string argument)
{
    var parts = argument.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2)
    {
        throw new InvalidOperationException("Usage: find CUSTOMERS Bremen");
    }

    return module.BuildSearchScreen(sessionState.Database, parts[0], parts[1]);
}

static ModuleScreen HandleDatabaseAppend(
    DatabaseModule module,
    AppSessionState sessionState,
    string argument,
    out bool persistState)
{
    var parts = argument.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2)
    {
        throw new InvalidOperationException("Usage: append CUSTOMERS Id=C-1004;Company=...;City=...;Tier=A");
    }

    persistState = true;
    return module.AppendRecord(sessionState.Database, parts[0], parts[1]);
}

static ModuleScreen HandleDatabaseUpdate(
    DatabaseModule module,
    AppSessionState sessionState,
    string argument,
    out bool persistState)
{
    var parts = argument.Split(' ', 4, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 4)
    {
        throw new InvalidOperationException("Usage: update CUSTOMERS C-1001 City Hannover");
    }

    persistState = true;
    return module.UpdateRecord(sessionState.Database, parts[0], parts[1], parts[2], parts[3]);
}

static ModuleScreen HandleDatabaseDelete(
    DatabaseModule module,
    AppSessionState sessionState,
    string argument,
    out bool persistState)
{
    var parts = argument.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2)
    {
        throw new InvalidOperationException("Usage: delete CUSTOMERS C-1004");
    }

    persistState = true;
    return module.DeleteRecord(sessionState.Database, parts[0], parts[1]);
}

static ModuleScreen HandleWordNew(
    WordProcessingModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    out bool persistState)
{
    sessionState.Word.Title = $"{workspace.Name} Letter Draft";
    sessionState.Word.Template = "Executive Letter";
    sessionState.Word.Lines =
    [
        "Dear customer,",
        "thank you for your continued business.",
        "Kind regards,",
        workspace.Owner
    ];
    persistState = true;
    return module.BuildNewLetterScreen(workspace, sessionState.Word);
}

static ModuleScreen HandleWordType(
    WordProcessingModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string text,
    out bool persistState)
{
    sessionState.Word.Lines.Add(text);
    persistState = true;
    return module.BuildNewLetterScreen(workspace, sessionState.Word);
}

static ModuleScreen HandleWordTitle(
    WordProcessingModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string title,
    out bool persistState)
{
    sessionState.Word.Title = title;
    persistState = true;
    return module.BuildNewLetterScreen(workspace, sessionState.Word);
}

static ModuleScreen HandleWordDeleteLine(
    WordProcessingModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    out bool persistState)
{
    sessionState.Word.DeleteLastLine();
    persistState = true;
    return module.BuildNewLetterScreen(workspace, sessionState.Word);
}

static ModuleScreen HandleMailTo(
    MailModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string recipient,
    out bool persistState)
{
    sessionState.Mail.To = recipient;
    persistState = true;
    return module.BuildComposeScreen(workspace, sessionState.Mail);
}

static ModuleScreen HandleMailSubject(
    MailModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string subject,
    out bool persistState)
{
    sessionState.Mail.Subject = subject;
    persistState = true;
    return module.BuildComposeScreen(workspace, sessionState.Mail);
}

static ModuleScreen HandleMailBody(
    MailModule module,
    OfficeWorkspace workspace,
    AppSessionState sessionState,
    string body,
    out bool persistState)
{
    sessionState.Mail.Body = body;
    persistState = true;
    return module.BuildComposeScreen(workspace, sessionState.Mail);
}

static ModuleScreen HandleMailSend(
    MailModule module,
    AppSessionState sessionState,
    out bool persistState)
{
    sessionState.Mail.SentItems.Add(new MailSentItemState
    {
        To = sessionState.Mail.To,
        Subject = sessionState.Mail.Subject,
        Body = sessionState.Mail.Body,
        SentAt = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
    });
    sessionState.Mail.To = "SALES";
    sessionState.Mail.Subject = "Weekly pipeline update";
    sessionState.Mail.Body = "Please send the current pipeline figures before noon.";
    persistState = true;
    return module.BuildSentItemsScreen(sessionState.Mail);
}

static string Fit(string text, int width)
{
    if (text.Length > width)
    {
        return text[..width];
    }

    return text.PadRight(width);
}

static string GetMenuLabel(string moduleId)
{
    return moduleId.ToLowerInvariant() switch
    {
        "comm" => "COMM",
        "db" => "DATA",
        "mail" => "MAIL",
        "pro" => "PROG",
        "report" => "RPT",
        "sheet" => "CALC",
        "word" => "WORD",
        _ => moduleId.ToUpperInvariant()
    };
}

static IEnumerable<string> Wrap(string text, int width)
{
    if (string.IsNullOrEmpty(text))
    {
        yield return string.Empty;
        yield break;
    }

    var remaining = text;
    while (remaining.Length > width)
    {
        var split = remaining.LastIndexOf(' ', Math.Min(width, remaining.Length - 1), width);
        if (split <= 0)
        {
            split = width;
        }

        yield return remaining[..split].TrimEnd();
        remaining = remaining[split..].TrimStart();
    }

    yield return remaining;
}

static void SafeClear()
{
    if (Console.IsOutputRedirected)
    {
        return;
    }

    Console.BackgroundColor = ConsoleColor.DarkBlue;
    Console.ForegroundColor = ConsoleColor.White;
    Console.Clear();
    Console.ResetColor();
}

static bool IsQuit(string input)
{
    return string.Equals(input, "q", StringComparison.OrdinalIgnoreCase) ||
           string.Equals(input, "quit", StringComparison.OrdinalIgnoreCase) ||
           string.Equals(input, "exit", StringComparison.OrdinalIgnoreCase);
}

static string NormalizeCommandAlias(string input)
{
    return input.ToUpperInvariant() switch
    {
        "/F1" or "F1" => "menu",
        "/F2" or "F2" => "menu",
        "/F3" or "F3" => "home",
        "/F10" or "F10" => "q",
        _ => input
    };
}

static string ReadCommand(IOfficeModule activeModule)
{
    if (Console.IsInputRedirected || Console.IsOutputRedirected)
    {
        Console.Write("Command> ");
        return Console.ReadLine()?.Trim() ?? string.Empty;
    }

    var buffer = new List<char>();
    Console.Write("Command> ");

    while (true)
    {
        var key = Console.ReadKey(intercept: true);

        switch (key.Key)
        {
            case ConsoleKey.Enter:
                Console.WriteLine();
                return new string(buffer.ToArray()).Trim();
            case ConsoleKey.Backspace:
                if (buffer.Count > 0)
                {
                    buffer.RemoveAt(buffer.Count - 1);
                    Console.Write("\b \b");
                }
                break;
            case ConsoleKey.F1:
                Console.WriteLine("menu");
                return "menu";
            case ConsoleKey.F2:
                Console.WriteLine("menu");
                return "menu";
            case ConsoleKey.F3:
                Console.WriteLine("home");
                return "home";
            case ConsoleKey.F5:
                Console.WriteLine("open ");
                return "open ";
            case ConsoleKey.F6:
                Console.WriteLine("edit ");
                return "edit ";
            case ConsoleKey.F7:
                Console.WriteLine("run ");
                return "run ";
            case ConsoleKey.F10:
                Console.WriteLine("q");
                return "q";
            case ConsoleKey.LeftArrow:
                Console.WriteLine();
                return "__NAV_LEFT__";
            case ConsoleKey.RightArrow:
                Console.WriteLine();
                return "__NAV_RIGHT__";
            default:
                if (!char.IsControl(key.KeyChar))
                {
                    buffer.Add(key.KeyChar);
                    Console.Write(key.KeyChar);
                }
                break;
        }
    }
}

static IOfficeModule CycleModule(OfficeSuite suite, IOfficeModule activeModule, int delta)
{
    var currentIndex = suite.Modules
        .Select((module, index) => new { module, index })
        .First(entry => entry.module.Info.Id == activeModule.Info.Id)
        .index;

    var nextIndex = (currentIndex + delta + suite.Modules.Count) % suite.Modules.Count;
    return suite.Modules[nextIndex];
}
