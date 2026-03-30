namespace SpiOpenAccess.Core;

public sealed class SpreadsheetWorkspaceState
{
    public decimal Q1 { get; set; } = 182_500m;
    public decimal Q2 { get; set; } = 204_400m;
    public decimal Q3 { get; set; } = 198_225m;
    public decimal Q4 { get; set; } = 227_000m;
    public string ActiveCell { get; set; } = "A1";
    public Dictionary<string, decimal> Cells { get; set; } = new(StringComparer.OrdinalIgnoreCase)
    {
        ["A1"] = 120_000m,
        ["B1"] = 132_500m,
        ["C1"] = 141_750m,
        ["D1"] = 155_900m,
        ["A2"] = 118_400m,
        ["B2"] = 127_300m,
        ["C2"] = 136_200m,
        ["D2"] = 149_880m
    };

    public decimal[] ToArray() => [Q1, Q2, Q3, Q4];

    public void SetQuarter(string quarter, decimal value)
    {
        switch (quarter.Trim().ToUpperInvariant())
        {
            case "Q1":
                Q1 = value;
                break;
            case "Q2":
                Q2 = value;
                break;
            case "Q3":
                Q3 = value;
                break;
            case "Q4":
                Q4 = value;
                break;
            default:
                throw new KeyNotFoundException($"Unknown quarter '{quarter}'.");
        }
    }

    public void SelectCell(string cellAddress)
    {
        var normalized = NormalizeCell(cellAddress);
        ActiveCell = normalized;
        if (!Cells.ContainsKey(normalized))
        {
            Cells[normalized] = 0m;
        }
    }

    public void SetCell(string cellAddress, decimal value)
    {
        var normalized = NormalizeCell(cellAddress);
        Cells[normalized] = value;
        ActiveCell = normalized;
    }

    public decimal GetCell(string cellAddress)
    {
        var normalized = NormalizeCell(cellAddress);
        return Cells.TryGetValue(normalized, out var value) ? value : 0m;
    }

    public static string NormalizeCell(string cellAddress)
    {
        var candidate = cellAddress.Trim().ToUpperInvariant();
        if (candidate.Length < 2 || candidate.Length > 3)
        {
            throw new InvalidOperationException($"Invalid cell '{cellAddress}'.");
        }

        var column = candidate[0];
        if (column is < 'A' or > 'H')
        {
            throw new InvalidOperationException($"Invalid cell '{cellAddress}'.");
        }

        if (!int.TryParse(candidate[1..], out var row) || row < 1 || row > 99)
        {
            throw new InvalidOperationException($"Invalid cell '{cellAddress}'.");
        }

        return candidate;
    }
}

public sealed class WordProcessorWorkspaceState
{
    public string Title { get; set; } = "OPENACCESS Proposal";
    public string Template { get; set; } = "Executive Letter";
    public List<string> Lines { get; set; } =
    [
        "Dear customer,",
        "please find the current commercial overview attached.",
        "Regards,",
        "ADMIN"
    ];
    public int CursorLine { get; set; } = 1;

    public void DeleteLastLine()
    {
        if (Lines.Count > 0)
        {
            Lines.RemoveAt(Lines.Count - 1);
        }
    }

    public int GetNormalizedCursorLine()
    {
        if (Lines.Count == 0)
        {
            Lines.Add(string.Empty);
        }

        if (CursorLine < 1)
        {
            CursorLine = 1;
        }
        else if (CursorLine > Lines.Count)
        {
            CursorLine = Lines.Count;
        }

        return CursorLine;
    }

    public void MoveCursor(int delta)
    {
        CursorLine = GetNormalizedCursorLine() + delta;
        GetNormalizedCursorLine();
    }

    public void InsertAtCursor(string text)
    {
        var index = GetNormalizedCursorLine() - 1;
        Lines.Insert(index, text);
        CursorLine = index + 1;
    }

    public void ReplaceAtCursor(string text)
    {
        var index = GetNormalizedCursorLine() - 1;
        Lines[index] = text;
    }

    public void DeleteAtCursor()
    {
        var index = GetNormalizedCursorLine() - 1;
        Lines.RemoveAt(index);
        if (Lines.Count == 0)
        {
            Lines.Add(string.Empty);
        }

        GetNormalizedCursorLine();
    }
}

public sealed class MailWorkspaceState
{
    public string To { get; set; } = "SALES";
    public string Subject { get; set; } = "Weekly pipeline update";
    public string Body { get; set; } = "Please send the current pipeline figures before noon.";
    public List<MailSentItemState> SentItems { get; set; } = [];
}

public sealed class MailSentItemState
{
    public string To { get; set; } = string.Empty;
    public string Subject { get; set; } = string.Empty;
    public string Body { get; set; } = string.Empty;
    public string SentAt { get; set; } = string.Empty;
}

public sealed class CommunicationsWorkspaceState
{
    public string CurrentTarget { get; set; } = "HQ";
    public bool IsConnected { get; set; }
    public string CaptureMode { get; set; } = "off";
    public string LastTransferFile { get; set; } = "orders.dat";
    public List<string> SessionLog { get; set; } =
    [
        "Profile loaded: ADMIN-OPS",
        "Last sync completed"
    ];
}

public sealed class ReportingWorkspaceState
{
    public string LastRunReport { get; set; } = "Aging";
    public string LastRunAt { get; set; } = string.Empty;
    public string ScheduledQueue { get; set; } = "Evening Finance";
    public string ActiveLayout { get; set; } = "Revenue-City";
    public List<string> OutputHistory { get; set; } =
    [
        "Aging-2026-03-28.lst",
        "Revenue-City.prn"
    ];
}

public sealed class ProgrammingWorkspaceState
{
    public string ProgramName { get; set; } = "sample.pro";
    public List<string> SourceLines { get; set; } =
    [
        "LET company = \"Nordwest Handel\"",
        "LET openInvoices = 14",
        "LET agingDays = 63",
        "PRINT company",
        "PRINT openInvoices + agingDays"
    ];
    public string LastCompileTarget { get; set; } = "nightly.pbc";
    public string LastRunAt { get; set; } = string.Empty;
}

public sealed class DatabaseWorkspaceState
{
    public List<DatabaseTableState> Tables { get; set; } = [];
    public string ActiveTableName { get; set; } = string.Empty;

    public DatabaseTableState GetTable(string tableName)
    {
        return Tables.FirstOrDefault(table => string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
            ?? throw new KeyNotFoundException($"Unknown table '{tableName}'.");
    }

    public DatabaseTableState GetActiveTable()
    {
        if (!string.IsNullOrWhiteSpace(ActiveTableName))
        {
            return GetTable(ActiveTableName);
        }

        if (Tables.Count == 0)
        {
            throw new KeyNotFoundException("No tables available.");
        }

        ActiveTableName = Tables[0].Name;
        return Tables[0];
    }

    public void SetActiveTable(string tableName)
    {
        var table = GetTable(tableName);
        ActiveTableName = table.Name;
    }
}

public sealed class DatabaseTableState
{
    public string Name { get; set; } = string.Empty;

    public List<Dictionary<string, string>> Records { get; set; } = [];
    public int CurrentIndex { get; set; }

    public int GetNormalizedCurrentIndex()
    {
        if (Records.Count == 0)
        {
            CurrentIndex = 0;
            return 0;
        }

        if (CurrentIndex < 0)
        {
            CurrentIndex = 0;
        }
        else if (CurrentIndex >= Records.Count)
        {
            CurrentIndex = Records.Count - 1;
        }

        return CurrentIndex;
    }

    public Dictionary<string, string>? GetCurrentRecord()
    {
        return Records.Count == 0 ? null : Records[GetNormalizedCurrentIndex()];
    }

    public int MoveNext()
    {
        if (Records.Count == 0)
        {
            CurrentIndex = 0;
            return CurrentIndex;
        }

        CurrentIndex = (GetNormalizedCurrentIndex() + 1) % Records.Count;
        return CurrentIndex;
    }

    public int MovePrevious()
    {
        if (Records.Count == 0)
        {
            CurrentIndex = 0;
            return CurrentIndex;
        }

        CurrentIndex = (GetNormalizedCurrentIndex() - 1 + Records.Count) % Records.Count;
        return CurrentIndex;
    }
}
