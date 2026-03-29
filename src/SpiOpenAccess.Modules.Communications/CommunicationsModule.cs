using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Communications;

public sealed class CommunicationsModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "comm",
        "Communications",
        "Terminalprofile, Dateitransfer und Skriptsteuerung.",
        "Connectivity",
        ["Terminal", "Transfer", "Profiles", "Automation"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        var content = new[]
        {
            $"Profile set      : {workspace.Owner}-OPS",
            "Connections      : COM1, COM2, TCP bridge",
            "Protocols        : XMODEM, YMODEM, ASCII capture",
            "Last session     : BBS mirror sync completed",
            "Log buffer       : 128 KB",
            "Automation       : Nightly upload + receive"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["dial HQ", "send orders.dat", "capture on"]);
    }
}
