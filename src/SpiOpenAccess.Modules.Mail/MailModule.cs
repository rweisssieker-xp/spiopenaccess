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
        var content = new List<string>
        {
            $"Mailbox          : {workspace.Owner}@{workspace.Name.ToLowerInvariant()}.lan",
            $"Unread           : {_inbox.Count(message => message.IsUnread)}",
            "Folders          : Inbox, Action, Archive, Outbox",
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

    private sealed record MessagePreview(string Id, string Subject, string Sender, bool IsUnread);
}
