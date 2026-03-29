using System.Globalization;
using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Programming;

public sealed class ProgrammingModule : IOfficeModule
{
    private const string SampleProgram = """
        LET company = "Nordwest Handel"
        LET openInvoices = 14
        LET agingDays = 63
        PRINT company
        PRINT openInvoices + agingDays
        """;

    public ModuleInfo Info { get; } = new(
        "pro",
        "Programming",
        "PRO-kompatible Skriptoberflaeche fuer Automatisierung und Datenlogik.",
        "Automation",
        ["Interpreter", "Variables", "Expressions", "Batch Jobs"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var runtime = RunProgram(SampleProgram);
        var content = new List<string>
        {
            $"Workspace         : {workspace.Name}",
            "Language          : PRO-inspired scripting runtime",
            $"Statements        : {runtime.StatementCount}",
            $"Variables         : {runtime.Variables.Count}",
            "Sample output     :"
        };
        content.AddRange(runtime.Output.Select(line => $"  {line}"));
        content.Add("Built-ins         : LET, PRINT, +, numeric and string literals");

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["run sample.pro", "list variables", "compile nightly.pro"]);
    }

    public ModuleScreen BuildRunScreen(string programName)
    {
        var result = RunProgram(SampleProgram);
        var content = new List<string>
        {
            $"Program          : {programName}",
            $"Statements       : {result.StatementCount}",
            $"Variables        : {result.Variables.Count}",
            "Output           :"
        };
        content.AddRange(result.Output.Select(line => $"  {line}"));

        return ModuleScreen.Create(
            $"Run {programName}",
            "Ausfuehrung eines PRO-inspirierten Programms.",
            content,
            ["list variables", "compile nightly.pro", "back"]);
    }

    public ModuleScreen BuildVariablesScreen()
    {
        var result = RunProgram(SampleProgram);
        var content = result.Variables
            .Select(pair => $"  {pair.Key,-16} {Convert.ToString(pair.Value, CultureInfo.InvariantCulture)}")
            .Prepend("Variable state    :")
            .ToArray();

        return ModuleScreen.Create(
            "Variable Watch",
            "Laufzeitstatus der Skriptvariablen.",
            content,
            ["run sample.pro", "compile nightly.pro", "back"]);
    }

    public ModuleScreen BuildCompileScreen(string programName)
    {
        return ModuleScreen.Create(
            $"Compile {programName}",
            "Vorbereitung eines Batch-Programms.",
            new[]
            {
                $"Source           : {programName}",
                "Passes           : lex, parse, bind",
                "Warnings         : 0",
                "Output           : nightly.pbc",
                "Status           : Compile successful"
            },
            ["run sample.pro", "list variables", "back"]);
    }

    public ProgramResult RunProgram(string source)
    {
        var variables = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        var output = new List<string>();
        var statements = 0;

        foreach (var rawLine in source.ReplaceLineEndings("\n").Split('\n', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            if (string.IsNullOrWhiteSpace(rawLine))
            {
                continue;
            }

            statements++;
            if (rawLine.StartsWith("LET ", StringComparison.OrdinalIgnoreCase))
            {
                var assignment = rawLine[4..];
                var parts = assignment.Split('=', 2, StringSplitOptions.TrimEntries);
                if (parts.Length != 2)
                {
                    throw new InvalidOperationException($"Invalid LET statement: {rawLine}");
                }

                variables[parts[0]] = Evaluate(parts[1], variables);
                continue;
            }

            if (rawLine.StartsWith("PRINT ", StringComparison.OrdinalIgnoreCase))
            {
                var value = Evaluate(rawLine[6..], variables);
                output.Add(Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty);
                continue;
            }

            throw new InvalidOperationException($"Unknown statement: {rawLine}");
        }

        return new ProgramResult(statements, variables, output);
    }

    private static object? Evaluate(string expression, IReadOnlyDictionary<string, object?> variables)
    {
        var parts = expression.Split('+', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length == 1)
        {
            return ResolveValue(parts[0], variables);
        }

        decimal numericResult = 0;
        foreach (var part in parts)
        {
            var value = ResolveValue(part, variables);
            numericResult += Convert.ToDecimal(value, CultureInfo.InvariantCulture);
        }

        return numericResult;
    }

    private static object? ResolveValue(string token, IReadOnlyDictionary<string, object?> variables)
    {
        if (token.StartsWith('"') && token.EndsWith('"') && token.Length >= 2)
        {
            return token[1..^1];
        }

        if (decimal.TryParse(token, NumberStyles.Number, CultureInfo.InvariantCulture, out var numericValue))
        {
            return numericValue;
        }

        if (variables.TryGetValue(token, out var variableValue))
        {
            return variableValue;
        }

        throw new KeyNotFoundException($"Unknown identifier: {token}");
    }
}

public sealed record ProgramResult(
    int StatementCount,
    IReadOnlyDictionary<string, object?> Variables,
    IReadOnlyList<string> Output);
