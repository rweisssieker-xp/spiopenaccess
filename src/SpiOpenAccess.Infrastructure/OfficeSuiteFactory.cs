using SpiOpenAccess.Core;
using SpiOpenAccess.Modules.Communications;
using SpiOpenAccess.Modules.Database;
using SpiOpenAccess.Modules.Mail;
using SpiOpenAccess.Modules.Programming;
using SpiOpenAccess.Modules.Reporting;
using SpiOpenAccess.Modules.Spreadsheet;
using SpiOpenAccess.Modules.WordProcessing;

namespace SpiOpenAccess.Infrastructure;

public static class OfficeSuiteFactory
{
    public static OfficeSuite CreateDefault()
    {
        var workspace = new OfficeWorkspace(
            "OPENACCESS",
            "ADMIN",
            new DateOnly(2026, 3, 28),
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                ["theme"] = "amber",
                ["shell"] = "text-mode",
                ["network"] = "enabled"
            });

        var modules = new IOfficeModule[]
        {
            new DatabaseModule(),
            new SpreadsheetModule(),
            new WordProcessingModule(),
            new CommunicationsModule(),
            new MailModule(),
            new ReportingModule(),
            new ProgrammingModule()
        };

        return new OfficeSuite("SPI Open Access Rebuild", "0.1.0", workspace, modules);
    }
}
