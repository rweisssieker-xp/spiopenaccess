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
        "Local mailboxes, folders, address book, and internal messaging.",
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
            $"Sent items       : {draft.SentItems.Count}",
            "Inbox preview    :"
        };
        content.AddRange(_inbox.Select(message =>
            $"  {(message.IsUnread ? "*" : " ")} {message.Id,-8} {message.Subject} [{message.Sender}]"));

        return ModuleScreen.Create(
            Info.DisplayName,
            Info.Summary,
            content,
            ["open OPS-142", "compose", "sent", "route rules"]);
    }

    public ModuleScreen BuildMessageScreen(string messageId)
    {
        var message = _inbox.FirstOrDefault(candidate =>
            string.Equals(candidate.Id, messageId, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown message '{messageId}'.");

        return ModuleScreen.Create(
            $"Message {message.Id}",
            "Message view for the local mailbox.",
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
            "Compose a new internal message.",
            new[]
            {
                $"From             : {workspace.Owner}@{workspace.Name.ToLowerInvariant()}.lan",
                $"To               : {draft.To}",
                "Cc               :",
                $"Subject          : {draft.Subject}",
                "Attachment       : none",
                $"Draft body       : {draft.Body}"
            },
            ["to <recipient>", "subject <text>", "body <text>", "send"]);
    }

    public ModuleScreen BuildSentItemsScreen(MailWorkspaceState draft)
    {
        var content = new List<string>
        {
            $"Sent count       : {draft.SentItems.Count}",
            "Outbox history    :"
        };
        content.AddRange(draft.SentItems
            .OrderByDescending(item => item.SentAt)
            .Select(item => $"  {item.SentAt,-19} {item.To,-12} {item.Subject}"));
        if (draft.SentItems.Count == 0)
        {
            content.Add("  <none>");
        }

        return ModuleScreen.Create(
            "Sent Items",
            "Recently sent internal messages.",
            content,
            ["compose", "back"]);
    }

    public ModuleScreen BuildRoutingRulesScreen()
    {
        return ModuleScreen.Create(
            "Routing Rules",
            "Automatic routing and folder rules.",
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
