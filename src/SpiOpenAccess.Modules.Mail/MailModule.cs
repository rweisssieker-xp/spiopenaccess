using SpiOpenAccess.Core;

namespace SpiOpenAccess.Modules.Mail;

public sealed class MailModule : IOfficeModule
{
    private readonly IReadOnlyList<MessagePreview> _inbox =
    [
        new("OPS-142", "Quarter close checklist", "Finance", true),
        new("OPS-143", "New lead import completed", "Sales", false),
        new("OPS-144", "Open invoices over 60 days", "Collections", true)
    ];

    public ModuleInfo Info { get; } = new(
        "mail",
        "Mail",
        "Lokale Mailboxen, Ordner, Adressbuch und interne Nachrichten.",
        "Messaging",
        ["Inbox", "Folders", "Address Book", "Routing"]);

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace)
    {
        return BuildHomeScreen(workspace, new MailWorkspaceState());
    }

    public ModuleScreen BuildHomeScreen(OfficeWorkspace workspace, MailWorkspaceState draft)
    {
        var content = new List<string>
        {
            $"Mailbox          : {workspace.Owner}@{workspace.Name.ToLowerInvariant()}.lan",
            $"Unread           : {_inbox.Count(message => message.IsUnread)}",
            "Folders          : Inbox, Action, Archive, Outbox",
            $"Draft to         : {draft.To}",
            $"Draft subject    : {draft.Subject}",
            "Inbox preview    :"
        };
        content.AddRange(_inbox.Select(message =>
            $"  {(message.IsUnread ? "*" : " ")} {message.Id,-8} {message.Subject} [{message.Sender}]"));

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["open OPS-142", "compose", "route rules"]);
    }

    public ModuleScreen BuildMessageScreen(string messageId)
    {
        var message = _inbox.FirstOrDefault(candidate =>
            string.Equals(candidate.Id, messageId, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown message '{messageId}'.");

        return ModuleScreen.Create(
            $"Message {message.Id}",
            "Nachrichtenansicht der lokalen Mailbox.",
            new[]
            {
                $"From             : {message.Sender}",
                $"Subject          : {message.Subject}",
                $"Status           : {(message.IsUnread ? "Unread" : "Read")}",
                "Folder           : Inbox",
                "Body             :",
                "  Action item review required before noon.",
                "  Please confirm by internal reply."
            },
            ["compose", "route rules", "back"]);
    }

    public ModuleScreen BuildComposeScreen(OfficeWorkspace workspace, MailWorkspaceState draft)
    {
        return ModuleScreen.Create(
            "Compose Message",
            "Neue interne Nachricht.",
            new[]
            {
                $"From             : {workspace.Owner}@{workspace.Name.ToLowerInvariant()}.lan",
                $"To               : {draft.To}",
                "Cc               :",
                $"Subject          : {draft.Subject}",
                "Attachment       : none",
                $"Draft body       : {draft.Body}"
            },
            ["to <recipient>", "subject <text>", "body <text>"]);
    }

    public ModuleScreen BuildRoutingRulesScreen()
    {
        return ModuleScreen.Create(
            "Routing Rules",
            "Automatische Verteilung und Ordnerregeln.",
            new[]
            {
                "Rule 1           : Finance -> Action",
                "Rule 2           : Subject contains 'invoice' -> Collections",
                "Rule 3           : Unread older than 3 days -> Escalation",
                "Default folder   : Inbox"
            },
            ["compose", "open OPS-142", "back"]);
    }

    private sealed record MessagePreview(string Id, string Subject, string Sender, bool IsUnread);
}
