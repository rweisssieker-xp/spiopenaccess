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
}
