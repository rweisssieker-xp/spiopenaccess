using SpiOpenAccess.Core;
using SpiOpenAccess.Infrastructure;
using SpiOpenAccess.Modules.Database;

const int FrameWidth = 80;
const int VisibleBodyLines = 17;
var suite = OfficeSuiteFactory.CreateDefault();

if (args.Length > 0)
{
    var selector = string.Join(' ', args);
    RenderSelectedModule(suite, selector);
    return;
}

RenderShell(suite);

static void RenderShell(OfficeSuite suite)
{
    var activeModule = suite.FindModule("db") ?? suite.Modules[0];
    var currentScreen = activeModule.BuildHomeScreen(suite.Workspace);
    var status = "F1 Menu  F2 Modules  F3 Home  F10 Exit";

    while (true)
    {
        RenderWorkspace(suite, activeModule, currentScreen, status);
        Console.Write("Command> ");
        var input = Console.ReadLine()?.Trim();
        if (string.IsNullOrWhiteSpace(input))
        {
            status = "No command entered.";
            continue;
        }

        input = NormalizeCommandAlias(input);

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
                currentScreen = activeModule.BuildHomeScreen(suite.Workspace);
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
            currentScreen = activeModule.BuildHomeScreen(suite.Workspace);
            status = $"Switched to {activeModule.Info.DisplayName}.";
            continue;
        }

        if (TryBuildCommandScreen(activeModule, suite.Workspace, input, out var commandScreen, out var error))
        {
            currentScreen = commandScreen!;
            status = $"Executed: {input}";
            continue;
        }

        status = error ?? $"Unknown command: {input}";
    }
}

static void RenderSelectedModule(OfficeSuite suite, string selector)
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
        module.BuildHomeScreen(suite.Workspace),
        "Direct module view.");
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
    Console.WriteLine(Fit(" F1=Menu  F2=Modules  F3=Home  F5=Open  F6=Edit  F7=Run  F10=Exit ", FrameWidth));
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
    string input,
    out ModuleScreen? screen,
    out string? error)
{
    screen = null;
    error = null;

    if (string.Equals(input, "back", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(input, "home", StringComparison.OrdinalIgnoreCase))
    {
        screen = activeModule.BuildHomeScreen(workspace);
        return true;
    }

    if (activeModule is not DatabaseModule databaseModule)
    {
        error = $"Command not supported in {activeModule.Info.DisplayName}.";
        return false;
    }

    var commandParts = input.Split(' ', 2, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (commandParts.Length != 2)
    {
        return false;
    }

    try
    {
        screen = commandParts[0].ToLowerInvariant() switch
        {
            "open" => databaseModule.BuildTableScreen(commandParts[1]),
            "edit" => databaseModule.BuildFormScreen(commandParts[1]),
            "run" => databaseModule.BuildReportScreen(commandParts[1]),
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
