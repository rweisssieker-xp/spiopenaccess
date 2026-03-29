using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.WordProcessing;

public sealed class WordProcessingModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "word",
        "Word Processor",
        "Dokumente, Formatvorlagen, Serienbriefe und Drucklayout.",
        "Documents",
        ["Documents", "Styles", "Merge Fields", "Print Layout"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var content = new[]
        {
            $"Document         : {workspace.Name} Proposal",
            "Template         : Executive Letter",
            "Pages            : 4",
            "Styles           : Body, Heading 1, Heading 2, Signature",
            "Merge fields     : Company, Contact, DueDate, NetAmount",
            "Print settings   : A4 portrait, 2.5 cm margins",
            "Review           : Spellcheck dictionary loaded"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["new letter", "merge customers", "preview page 1"]);
    }

    public ModuleScreen BuildNewLetterScreen(OfficeWorkspace workspace)
    {
        return ModuleScreen.Create(
            "New Letter",
            "Neues Dokument auf Basis einer Vorlage.",
            new[]
            {
                $"Document         : {workspace.Name} Letter Draft",
                "Template         : Executive Letter",
                "Cursor           : Page 1, Line 1, Col 1",
                "Default style    : Body",
                "Insert blocks    : Address, Subject, Greeting, Signature",
                "Status           : Draft created"
            },
            ["merge customers", "preview page 1", "back"]);
    }

    public ModuleScreen BuildMergeScreen()
    {
        return ModuleScreen.Create(
            "Mail Merge",
            "Serienbriefvorschau mit Datenbankfeldern.",
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

    public ModuleScreen BuildPreviewScreen(string page)
    {
        return ModuleScreen.Create(
            "Print Preview",
            "Seitenvorschau des aktiven Dokuments.",
            new[]
            {
                $"Page             : {page}",
                "Zoom             : Full page",
                "Header           : Executive Letter",
                "Body lines       : 31",
                "Footer           : Confidential"
            },
            ["new letter", "merge customers", "back"]);
    }
}
