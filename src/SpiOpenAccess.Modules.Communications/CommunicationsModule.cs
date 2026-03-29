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

    public ModuleScreen BuildDialScreen(string target)
    {
        return ModuleScreen.Create(
            "Dial Session",
            "Aufbau einer Kommunikationssitzung.",
            new[]
            {
                $"Target           : {target}",
                "Port             : COM1",
                "Baud             : 9600",
                "Parity           : None",
                "Handshake        : XON/XOFF",
                "Status           : Carrier detected"
            },
            ["send orders.dat", "capture on", "back"]);
    }

    public ModuleScreen BuildSendScreen(string fileName)
    {
        return ModuleScreen.Create(
            "File Transfer",
            "Dateiuebertragung ueber aktives Kommunikationsprofil.",
            new[]
            {
                $"Filename         : {fileName}",
                "Protocol         : YMODEM",
                "Blocks sent      : 18",
                "Retries          : 0",
                "Status           : Transfer complete"
            },
            ["dial HQ", "capture on", "back"]);
    }

    public ModuleScreen BuildCaptureScreen(string mode)
    {
        return ModuleScreen.Create(
            "Capture Buffer",
            "Protokolliert eingehende Daten in eine Mitschnittdatei.",
            new[]
            {
                $"Capture mode     : {mode}",
                "Output file      : HQ-SESSION.LOG",
                "Size limit       : 128 KB",
                "Rotation         : Disabled",
                "Status           : Capture armed"
            },
            ["dial HQ", "send orders.dat", "back"]);
    }
}
