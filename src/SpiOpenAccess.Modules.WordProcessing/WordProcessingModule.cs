using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.WordProcessing;

public sealed class WordProcessingModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "word",
        "Word Processor",
        "Documents, styles, mail merge, and print layout.",
        "Documents",
        ["Documents", "Styles", "Merge Fields", "Print Layout"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        return BuildHomeScreen(workspace, new WordProcessorWorkspaceState());
    }

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace, WordProcessorWorkspaceState state)
    {
        var content = new[]
        {
            $"Document         : {state.Title}",
            $"Template         : {state.Template}",
            $"Pages            : {Math.Max(1, (state.Lines.Count + 24) / 25)}",
            "Styles           : Body, Heading 1, Heading 2, Signature",
            "Merge fields     : Company, Contact, DueDate, NetAmount",
            "Print settings   : A4 portrait, 2.5 cm margins",
            $"Review           : {state.Lines.Count} content lines loaded"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["new letter", "merge customers", "preview page 1"]);
    }

    public ModuleScreen BuildNewLetterScreen(OfficeWorkspace workspace, WordProcessorWorkspaceState state)
    {
        return ModuleScreen.Create(
            "New Letter",
            "Create a new document from a template.",
            new[]
            {
                $"Document         : {state.Title}",
                $"Template         : {state.Template}",
                $"Cursor           : Page 1, Line {Math.Max(1, state.Lines.Count)}, Col 1",
                "Default style    : Body",
                "Insert blocks    : Address, Subject, Greeting, Signature",
                $"Status           : {state.Lines.Count} lines in draft"
            },
            ["type <text>", "title <text>", "preview page 1"]);
    }

    public ModuleScreen BuildMergeScreen()
    {
        return ModuleScreen.Create(
            "Mail Merge",
            "Mail merge preview using database fields.",
            new[]
            {
                "Source form      : CUSTOMER_CARD",
                "Records queued   : 3",
                "Fields mapped    : Company, City, DueDate, NetAmount",
                "Output mode      : Print + archive",
                "Preview target   : Nordwest Handel"
            },
            ["new letter", "preview page 1", "back"]);
    }

    public ModuleScreen BuildPreviewScreen(string page, WordProcessorWorkspaceState state)
    {
        var previewLines = state.Lines.Take(5).Select(line => $"  {line}").ToArray();
        return ModuleScreen.Create(
            "Print Preview",
            "Page preview of the active document.",
            new[]
            {
                $"Page             : {page}",
                "Zoom             : Full page",
                $"Header           : {state.Template}",
                $"Body lines       : {state.Lines.Count}",
                "Footer           : Confidential"
            }.Concat(previewLines),
            ["new letter", "merge customers", "back"]);
    }
}
