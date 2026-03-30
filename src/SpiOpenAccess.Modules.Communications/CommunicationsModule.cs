using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Communications;

public sealed class CommunicationsModule : IOfficeModule
{
    public ModuleInfo Info { get; } = new(
        "comm",
        "Communications",
        "Terminal profiles, file transfer, and scripted communication.",
        "Connectivity",
        ["Terminal", "Transfer", "Profiles", "Automation"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        return BuildHomeScreen(workspace, new CommunicationsWorkspaceState());
    }

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace, CommunicationsWorkspaceState state)
    {
        var content = new[]
        {
            $"Profile set      : {workspace.Owner}-OPS",
            "Connections      : COM1, COM2, TCP bridge",
            "Protocols        : XMODEM, YMODEM, ASCII capture",
            $"Current target   : {state.CurrentTarget}",
            $"Connection state : {(state.IsConnected ? "Connected" : "Idle")}",
            $"Capture mode     : {state.CaptureMode}",
            $"Last transfer    : {state.LastTransferFile}",
            "Log buffer       : 128 KB",
            $"Session log      : {(state.SessionLog.Count == 0 ? "<empty>" : state.SessionLog[^1])}"
        };

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["dial HQ", "send orders.dat", "capture on"]);
    }

    public ModuleScreen BuildDialScreen(string target)
    {
        return BuildDialScreen(new CommunicationsWorkspaceState { CurrentTarget = target, IsConnected = true }, target);
    }

    public ModuleScreen BuildDialScreen(CommunicationsWorkspaceState state, string target)
    {
        return ModuleScreen.Create(
            "Dial Session",
            "Communication session setup.",
            new[]
            {
                $"Target           : {target}",
                "Port             : COM1",
                "Baud             : 9600",
                "Parity           : None",
                "Handshake        : XON/XOFF",
                $"Status           : {(state.IsConnected ? "Carrier detected" : "Idle")}"
            },
            ["send orders.dat", "capture on", "back"]);
    }

    public ModuleScreen BuildSendScreen(string fileName)
    {
        return BuildSendScreen(new CommunicationsWorkspaceState { LastTransferFile = fileName, IsConnected = true }, fileName);
    }

    public ModuleScreen BuildSendScreen(CommunicationsWorkspaceState state, string fileName)
    {
        return ModuleScreen.Create(
            "File Transfer",
            "File transfer over the active communications profile.",
            new[]
            {
                $"Filename         : {fileName}",
                "Protocol         : YMODEM",
                "Blocks sent      : 18",
                "Retries          : 0",
                $"Status           : {(state.IsConnected ? "Transfer complete" : "No carrier")}"
            },
            ["dial HQ", "capture on", "back"]);
    }

    public ModuleScreen BuildCaptureScreen(string mode)
    {
        return BuildCaptureScreen(new CommunicationsWorkspaceState { CaptureMode = mode }, mode);
    }

    public ModuleScreen BuildCaptureScreen(CommunicationsWorkspaceState state, string mode)
    {
        return ModuleScreen.Create(
            "Capture Buffer",
            "Capture incoming data into a session log.",
            new[]
            {
                $"Capture mode     : {mode}",
                "Output file      : HQ-SESSION.LOG",
                "Size limit       : 128 KB",
                "Rotation         : Disabled",
                $"Status           : {(string.Equals(state.CaptureMode, "on", StringComparison.OrdinalIgnoreCase) ? "Capture armed" : "Capture stopped")}"
            },
            ["dial HQ", "send orders.dat", "back"]);
    }
}
